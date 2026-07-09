const { DateTime } = require('luxon');
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
      clientId: secrets.GRAPH_CLIENT_ID,
      clientSecret: secrets.GRAPH_CLIENT_SECRET,
      authority: `https://login.microsoftonline.com/${secrets.GRAPH_TENANT_ID}`,
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

async function graphRequest(method, url, body, extraHeaders = null) {
  const maxAttempts = 5;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const token = await getGraphAccessToken();
      const response = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          ...(extraHeaders || {}),
        },
        body: body ? JSON.stringify(body) : undefined,
      });

      const text = await response.text();
      let data;
      try {
        data = text ? JSON.parse(text) : {};
      } catch {
        data = { raw: text };
      }

      if (!response.ok) {
        const retryable = response.status === 429 || (response.status >= 500 && response.status <= 599);
        if (retryable && attempt < maxAttempts) {
          await wait(getBackoffMs(attempt));
          continue;
        }
        throw new Error(`Graph ${method} failed ${response.status}: ${JSON.stringify(data)}`);
      }

      return data;
    } catch (error) {
      lastError = error;
      const msg = String(error?.message || '').toLowerCase();
      const retryable = msg.includes('fetch') || msg.includes('network') || msg.includes('timeout');
      if (retryable && attempt < maxAttempts) {
        await wait(getBackoffMs(attempt));
        continue;
      }
      throw error;
    }
  }

  throw lastError;
}

function getBackoffMs(attempt) {
  const jitter = Math.floor(Math.random() * 150);
  return Math.min(8_000, 300 * (2 ** (attempt - 1)) + jitter);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getCalendarWindowEvents(userUpn, pastDays, futureDays) {
  const now = DateTime.utc();
  const startDateTime = now.minus({ days: pastDays }).toISO();
  const endDateTime = now.plus({ days: futureDays }).toISO();
  return getCalendarEventsBetween(userUpn, startDateTime, endDateTime);
}

async function getCalendarEventsBetween(userUpn, startDateTime, endDateTime) {
  const selectFields = [
    'id',
    'iCalUId',
    'subject',
    'start',
    'end',
    'isAllDay',
    'showAs',
    'location',
    'attendees',
    'bodyPreview',
    'lastModifiedDateTime',
    'sensitivity',
  ];

  const params = new URLSearchParams({
    startDateTime: String(startDateTime),
    endDateTime: String(endDateTime),
    $select: selectFields.join(','),
    $top: '100',
  });

  let url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userUpn)}/calendarView?${params.toString()}`;
  const events = [];

  while (url) {
    const response = await graphRequest('GET', url);
    events.push(...(response.value || []));
    url = response['@odata.nextLink'] || null;
  }

  return events;
}

async function getDeltaEvents(userUpn, deltaLink = null, options = {}) {
  const {
    pastDays = 30,
    futureDays = 90,
  } = options;
  const baseUrl = 'https://graph.microsoft.com/v1.0';
  let url;

  if (deltaLink) {
    url = deltaLink;
  } else {
    const now = DateTime.utc();
    const startDateTime = now.minus({ days: pastDays }).toISO();
    const endDateTime = now.plus({ days: futureDays }).toISO();
    const params = new URLSearchParams({
      startDateTime,
      endDateTime,
      $select: 'subject,start,end,showAs,location,id,iCalUId,sensitivity,bodyPreview,isAllDay,lastModifiedDateTime',
    });
    url = `${baseUrl}/users/${encodeURIComponent(userUpn)}/calendarView/delta?${params.toString()}`;
  }

  const allEvents = [];
  // NOTE: Graph rejects $top for calendarView/delta; page size is controlled by Prefer: odata.maxpagesize.
  let response = await graphRequest('GET', url, null, { Prefer: 'odata.maxpagesize=50' });
  allEvents.push(...(response.value || []));

  while (response['@odata.nextLink']) {
    response = await graphRequest('GET', response['@odata.nextLink'], null, { Prefer: 'odata.maxpagesize=50' });
    allEvents.push(...(response.value || []));
  }

  return {
    events: allEvents,
    nextDeltaLink: response['@odata.deltaLink'] || null,
  };
}

function getSubscriptionResource(userUpn) {
  return `/users/${userUpn}/events`;
}

async function findSubscriptionByResource(resource, clientState) {
  const baseUrl = 'https://graph.microsoft.com/v1.0/subscriptions';
  // Some tenants reject $top on this endpoint.
  let url = baseUrl;

  while (url) {
    const response = await graphRequest('GET', url);
    const match = (response.value || []).find((subscription) => {
      const stateMatches = clientState ? subscription.clientState === clientState : true;
      return subscription.resource === resource && stateMatches;
    });
    if (match) {
      return match;
    }
    url = response['@odata.nextLink'] || null;
  }

  return null;
}

async function createSubscription(notificationUrl, clientState, userUpn) {
  const expirationDateTime = DateTime.utc().plus({ hours: 48 }).toISO();
  const resource = getSubscriptionResource(userUpn);

  return graphRequest('POST', 'https://graph.microsoft.com/v1.0/subscriptions', {
    changeType: 'created,updated,deleted',
    notificationUrl,
    resource,
    expirationDateTime,
    clientState,
  });
}

async function renewSubscription(subscriptionId) {
  const expirationDateTime = DateTime.utc().plus({ hours: 48 }).toISO();
  return graphRequest('PATCH', `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`, {
    expirationDateTime,
  });
}

async function createOrRenewSubscription(userUpn, notificationUrl, clientState) {
  const resource = getSubscriptionResource(userUpn);
  const existingSubscription = await findSubscriptionByResource(resource, clientState);

  if (existingSubscription) {
    return renewSubscription(existingSubscription.id);
  }

  return createSubscription(notificationUrl, clientState, userUpn);
}

module.exports = {
  getCalendarWindowEvents,
  getCalendarEventsBetween,
  getDeltaEvents,
  createSubscription,
  createOrRenewSubscription,
};
