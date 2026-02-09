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

async function graphRequest(method, url, body) {
  const token = await getGraphAccessToken();
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
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
    throw new Error(`Graph ${method} failed ${response.status}: ${JSON.stringify(data)}`);
  }

  return data;
}

async function getCalendarWindowEvents(userUpn, pastDays, futureDays) {
  const now = DateTime.utc();
  const startDateTime = now.minus({ days: pastDays }).toISO();
  const endDateTime = now.plus({ days: futureDays }).toISO();
  const selectFields = [
    'id',
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
    startDateTime,
    endDateTime,
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

async function getDeltaEvents(userUpn, deltaLink = null) {
  const baseUrl = 'https://graph.microsoft.com/v1.0';
  let url;

  if (deltaLink) {
    url = deltaLink;
  } else {
    const now = DateTime.utc();
    const startDateTime = now.minus({ days: 30 }).toISO();
    const endDateTime = now.plus({ days: 90 }).toISO();
    const params = new URLSearchParams({
      startDateTime,
      endDateTime,
      $select: 'subject,start,end,showAs,location,id,sensitivity,bodyPreview,isAllDay,lastModifiedDateTime',
      $top: '50',
    });
    url = `${baseUrl}/users/${encodeURIComponent(userUpn)}/calendarView/delta?${params.toString()}`;
  }

  const allEvents = [];
  let response = await graphRequest('GET', url);
  allEvents.push(...(response.value || []));

  while (response['@odata.nextLink']) {
    response = await graphRequest('GET', response['@odata.nextLink']);
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
  let url = `${baseUrl}?$top=100`;

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
  getDeltaEvents,
  createSubscription,
  createOrRenewSubscription,
};
