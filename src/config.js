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
