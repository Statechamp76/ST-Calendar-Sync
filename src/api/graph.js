const { ConfidentialClientApplication } = require('@azure/msal-node');
const { getSecrets } = require('../utils/secrets');

let msalClient;

async function getMsalClient() {
  if (msalClient) return msalClient;

  const secrets = await getSecrets([
    'GRAPH_CLIENT_ID',
    'GRAPH_CLIENT_SECRET',
    'GRAPH_TENANT_ID',
  ]);

  msalClient = new ConfidentialClientApplication({
    auth: {
      clientId: secrets.GRAPH_CLIENT_ID.trim(),
      clientSecret: secrets.GRAPH_CLIENT_SECRET.trim(),
      authority: `https://login.microsoftonline.com/${secrets.GRAPH_TENANT_ID.trim()}`,
    },
  });

  return msalClient;
}

async function getGraphAccessToken() {
  const client = await getMsalClient();
  const tokenResponse = await client.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return tokenResponse.accessToken;
}

async function graphGet(url) {
  const token = await getGraphAccessToken();
  const res = await fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  });
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch { data = { raw: text }; }
  if (!res.ok) {
    throw new Error(`Graph GET failed ${res.status}: ${JSON.stringify(data)}`);
  }
  return data;
}

/**
 * Retrieves delta changes for a user's calendarView
 */
async function getDeltaEvents(userUpn, deltaLink = null) {
  // Rolling 90-day window end = now + 90 days? (you said rolling 90; most do past 90. keep past 90 here.)
  const ninetyDaysAgo = new Date();
  ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
  const now = new Date();

  const baseUrl = 'https://graph.microsoft.com/v1.0';

  let url;
  if (deltaLink) {
    // deltaLink is a full URL from Graph
    url = deltaLink;
  } else {
    const params = new URLSearchParams({
      startDateTime: ninetyDaysAgo.toISOString(),
      endDateTime: now.toISOString(),
     '$select': 'subject,start,end,showAs,location,id,sensitivity,bodyPreview,isAllDay',
      '$top': '50',
    });
    url = `${baseUrl}/users/${encodeURIComponent(userUpn)}/calendarView/delta?${params.toString()}`;
  }

  let allEvents = [];
  let response = await graphGet(url);

  allEvents = allEvents.concat(response.value || []);

  // paginate if needed
  while (response['@odata.nextLink']) {
    response = await graphGet(response['@odata.nextLink']);
    allEvents = allEvents.concat(response.value || []);
  }

  return {
    events: allEvents,
    nextDeltaLink: response['@odata.deltaLink'] || null,
  };
}

/**
 * Create subscription (optional now; you can add later once delta works)
 */
async function createSubscription(notificationUrl, clientState) {
  const token = await getGraphAccessToken();
  const expirationDateTime = new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString(); // +2 days

  const res = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      changeType: 'created,updated,deleted',
      notificationUrl,
      resource: '/users/{id}/events', // NOTE: for per-user you’ll create per user; keep simple for now
      expirationDateTime,
      clientState,
    }),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(`Create subscription failed ${res.status}: ${JSON.stringify(data)}`);
  return data;
}

module.exports = {
  getDeltaEvents,
  createSubscription,
};


