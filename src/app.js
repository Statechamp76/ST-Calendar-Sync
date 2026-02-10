const express = require('express');
const { PubSub } = require('@google-cloud/pubsub');
const syncService = require('./services/sync');
const { requireOidcAuth } = require('./middleware/auth');
const { notifyFailure } = require('./services/alerts');
const { getSecrets } = require('./utils/secrets');
const { loadConfig } = require('./config');
const cleanupService = require('./services/cleanup');

const app = express();
app.use(express.json()); // Middleware to parse JSON bodies

const config = loadConfig();

// --- Endpoints ---

// Webhook receiver from Microsoft Graph
app.post('/graph/notifications', async (req, res) => {
    // Microsoft Graph webhook validation handshake
    const validationToken = req.query.validationToken;
    if (validationToken) {
        console.log('Received Graph webhook validation request.');
        res.status(200).send(validationToken);
        return;
    }

    // Process actual notifications
    // A robust implementation would verify the origin of the request (e.g., using clientState)
    // and handle decryption of encrypted notifications if configured.
    try {
        if (config.maintenanceMode) {
            console.warn('maintenance_mode: ignoring graph notifications publish');
            res.status(202).send();
            return;
        }

        if (!req.body || !Array.isArray(req.body.value)) {
            res.status(400).send('Bad Request: Invalid notification payload.');
            return;
        }

        const secrets = await getSecrets(['GRAPH_CLIENT_STATE']);
        const expectedClientState = secrets.GRAPH_CLIENT_STATE;
        const pubsub = new PubSub();
        const topicName = config.pubsubTopic; // Must exist in GCP

        // The notification body from Graph can contain multiple notifications
        for (const notification of req.body.value) {
            if (!notification || notification.clientState !== expectedClientState) {
                console.warn('Ignoring Graph notification with invalid clientState.');
                continue;
            }

            // Extract user identifier (UPN) from the resource path.
            // Example resource: 'users/someone@example.com/events/...'
            const resource = String(notification.resource || '');
            const upnMatch = resource.match(/users\/([^/]+)\/events/i);
            const userUpn = upnMatch ? decodeURIComponent(upnMatch[1]) : null;

            if (userUpn) {
                const messageId = await pubsub.topic(topicName).publishMessage({ json: { upn: userUpn } });
                console.log(`Published message ${messageId} for UPN: ${userUpn}`);
            } else {
                console.warn('Could not extract UPN from Graph notification resource:', resource);
            }
        }
        res.status(202).send(); // Accepted for processing
    } catch (error) {
        console.error('Error processing Graph webhook notification:', error);
        res.status(500).send('Internal Server Error');
    }
});

// Worker endpoint triggered by Pub/Sub push subscription
app.post('/sync/user', requireOidcAuth, async (req, res) => {
    if (config.maintenanceMode) {
        console.warn('maintenance_mode: ignoring /sync/user');
        res.status(204).send();
        return;
    }

    // Pub/Sub push messages arrive in the request body as a JSON object
    // containing a 'message' field which has the base64 encoded data.
    if (!req.body || !req.body.message || !req.body.message.data) {
        console.error('Invalid Pub/Sub message format received.');
        res.status(400).send('Bad Request: Invalid message format.');
        return;
    }

    try {
        const messageData = Buffer.from(req.body.message.data, 'base64').toString('utf-8');
        let message;
        try {
            message = JSON.parse(messageData);
        } catch {
            // Be tolerant of incorrectly-escaped JSON (or plain strings) so Pub/Sub doesn't retry forever.
            try {
                message = JSON.parse(messageData.replace(/\\"/g, '"'));
            } catch {
                message = messageData;
            }
        }

        let userUpn = typeof message === 'string' ? message : message.upn;

        // Additional tolerance: sometimes test publishes (or badly formatted payloads) arrive as `{upn:someone@...}`
        // which is not valid JSON. Extract the UPN if possible.
        if (typeof userUpn === 'string') {
            const trimmed = userUpn.trim();
            const m = trimmed.match(/upn[:=]\s*([^,}\s]+)/i);
            if (m && m[1]) {
                userUpn = m[1].trim();
            }
        }

        if (!userUpn) {
            console.error('Received Pub/Sub message with missing UPN:', message);
            // Ack the message to avoid endless retries on bad payloads.
            res.status(204).send();
            return;
        }
        
        console.log(`Received Pub/Sub message to sync user: ${userUpn}`);
        await syncService.runDeltaSyncForUser(userUpn); // Call the core sync logic
        
        res.status(204).send(); // Success, no content. Pub/Sub will acknowledge the message.
    } catch (error) {
  console.error('Error syncing user:', error);
  await notifyFailure('ST Calendar Sync: /sync/user failed', {
    message: error.message,
  });
  res.status(500).send('Internal Server Error');
}

});

// Endpoint for full nightly sync (Cloud Scheduler)
app.post('/sync/all', requireOidcAuth, async (req, res) => {
    try {
        console.log('Initiating full sync for all enabled users.');
        // This endpoint will iterate through TechMap and publish messages to Pub/Sub
        // for each enabled user, similar to the webhook, but for a full delta sync.
        await syncService.runFullSyncForAllUsers();
        res.status(202).send('Full sync initiated.');
    } catch (error) {
        console.error('Error initiating full sync:', error);
        res.status(500).send('Internal Server Error');
    }
});

// Endpoint for renewing Microsoft Graph subscriptions (Cloud Scheduler)
app.post('/graph/subscriptions/renew', requireOidcAuth, async (req, res) => {
    try {
        console.log('Attempting to renew Microsoft Graph subscriptions.');
        // This endpoint will retrieve existing subscriptions and renew them.
        await syncService.renewGraphSubscriptions();
        res.status(200).send('Subscription renewal process started.');
    } catch (error) {
        console.error('Error renewing Graph subscriptions:', error);
        res.status(500).send('Internal Server Error');
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.status(200).send('ok');
});

app.post('/run-sync', requireOidcAuth, async (req, res) => {
    try {
        const summary = await syncService.runSyncCycle();
        if (summary.errors && summary.errors.length > 0) {
            await notifyFailure('ST Calendar Sync: /run-sync completed with errors', {
                errorCount: summary.errors.length,
                sample: summary.errors.slice(0, 5),
            });
        }
        res.status(200).json(summary);
    } catch (error) {
        console.error('Run sync failed:', error);
        await notifyFailure('ST Calendar Sync: /run-sync failed', {
            message: error.message,
        });
        res.status(500).json({
            startedAt: new Date().toISOString(),
            finishedAt: new Date().toISOString(),
            calendarsProcessed: 0,
            eventsFetched: 0,
            eventsUpserted: 0,
            eventsSkipped: 0,
            errors: [{ message: error.message }],
        });
    }
});

// One-time backfill: pull last 30 days of calendarView for all enabled users and upsert only busy/OOF.
// Does not modify Outlook; it only creates/updates/deletes ServiceTitan non-job appointments + sheet mappings.
app.post('/backfill/last-30-days', requireOidcAuth, async (req, res) => {
    try {
        const summary = await syncService.runBackfillLast30DaysAllUsers();
        if (summary.errors && summary.errors.length > 0) {
            await notifyFailure('ST Calendar Sync: /backfill/last-30-days completed with errors', {
                errorCount: summary.errors.length,
                sample: summary.errors.slice(0, 5),
            });
        }
        res.status(200).json(summary);
    } catch (error) {
        console.error('Backfill last 30 days failed:', error);
        await notifyFailure('ST Calendar Sync: /backfill/last-30-days failed', {
            message: error.message,
        });
        res.status(500).json({
            startedAt: new Date().toISOString(),
            finishedAt: new Date().toISOString(),
            calendarsProcessed: 0,
            eventsFetched: 0,
            eventsUpserted: 0,
            eventsSkipped: 0,
            errors: [{ message: error.message }],
        });
    }
});

// One-time forward-looking backfill: pull next 90 days of calendarView for all enabled users and upsert only busy/OOF.
// Does not modify Outlook; it only creates/updates/deletes ServiceTitan non-job appointments + sheet mappings.
app.post('/backfill/next-90-days', requireOidcAuth, async (req, res) => {
    try {
        const summary = await syncService.runBackfillNext90DaysAllUsers();
        if (summary.errors && summary.errors.length > 0) {
            await notifyFailure('ST Calendar Sync: /backfill/next-90-days completed with errors', {
                errorCount: summary.errors.length,
                sample: summary.errors.slice(0, 5),
            });
        }
        res.status(200).json(summary);
    } catch (error) {
        console.error('Backfill next 90 days failed:', error);
        await notifyFailure('ST Calendar Sync: /backfill/next-90-days failed', {
            message: error.message,
        });
        res.status(500).json({
            startedAt: new Date().toISOString(),
            finishedAt: new Date().toISOString(),
            calendarsProcessed: 0,
            eventsFetched: 0,
            eventsUpserted: 0,
            eventsSkipped: 0,
            errors: [{ message: error.message }],
        });
    }
});

// One-time maintenance: deduplicate ServiceTitan non-job appointments created by this sync (Busy/Out of Office blockers)
// from the current week forward (default 90-day window), without touching Outlook.
app.post('/cleanup/deduplicate', requireOidcAuth, async (req, res) => {
    try {
        const body = req.body || {};
        const dryRun = body.dryRun !== false;
        const summary = await cleanupService.dedupeNonJobsThisWeekForward({
            startsOnOrAfter: body.startsOnOrAfter || null,
            startsOnOrBefore: body.startsOnOrBefore || null,
            dryRun,
        });
        console.log('cleanup.dedupe.complete', summary);
        res.status(200).json(summary);
    } catch (error) {
        console.error('Deduplicate failed:', error);
        await notifyFailure('ST Calendar Sync: /cleanup/deduplicate failed', {
            message: error.message,
        });
        res.status(500).json({ error: error.message });
    }
});

// Nuclear option: delete ALL ServiceTitan non-job appointments in the sync window for all enabled technicians,
// then clear EventMap + DeltaState so the next sync rebuilds from Graph.
// Never touches Outlook.
app.post('/cleanup/reset', requireOidcAuth, async (req, res) => {
    try {
        const body = req.body || {};
        const dryRun = body.dryRun !== false;
        const summary = await cleanupService.resetSyncState({
            startsOnOrAfter: body.startsOnOrAfter || null,
            startsOnOrBefore: body.startsOnOrBefore || null,
            dryRun,
            skipSheetsClear: body.skipSheetsClear === true,
        });
        console.log('cleanup.reset.complete', summary);
        res.status(200).json(summary);
    } catch (error) {
        console.error('Reset failed:', error);
        await notifyFailure('ST Calendar Sync: /cleanup/reset failed', {
            message: error.message,
        });
        res.status(500).json({ error: error.message });
    }
});

module.exports = app;
