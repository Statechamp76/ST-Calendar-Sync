const express = require('express');
const { PubSub } = require('@google-cloud/pubsub');
const syncService = require('./services/sync'); // Will be implemented later
const { requireOidcAuth } = require('./middleware/auth');
const { notifyFailure } = require('./services/alerts');

const app = express();
app.use(express.json()); // Middleware to parse JSON bodies

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
        if (!req.body || !Array.isArray(req.body.value)) {
            res.status(400).send('Bad Request: Invalid notification payload.');
            return;
        }

        const pubsub = new PubSub();
        const topicName = 'graph-notifications'; // Ensure this topic exists in GCP

        // The notification body from Graph can contain multiple notifications
        for (const notification of req.body.value) {
            // Extract user identifier (UPN) from the resource path.
            // Example resource: 'users/someone@example.com/events/...'
            const resource = notification.resource;
            const upnMatch = resource.match(/users\/(.*?)\/events/);
            const userUpn = upnMatch ? upnMatch[1] : null; 

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
app.post('/sync/user', async (req, res) => {
    // Pub/Sub push messages arrive in the request body as a JSON object
    // containing a 'message' field which has the base64 encoded data.
    if (!req.body || !req.body.message || !req.body.message.data) {
        console.error('Invalid Pub/Sub message format received.');
        res.status(400).send('Bad Request: Invalid message format.');
        return;
    }

    try {
        const messageData = Buffer.from(req.body.message.data, 'base64').toString('utf-8');
        const message = JSON.parse(messageData);
        const userUpn = message.upn;

        if (!userUpn) {
            console.error('Received Pub/Sub message with missing UPN:', message);
            res.status(400).send('Bad Request: Missing UPN in message.');
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
app.post('/sync/all', async (req, res) => {
    try {
        console.log('Initiating full sync for all enabled users.');
        // This endpoint will iterate through TechMap and publish messages to Pub/Sub
        // for each enabled user, similar to the webhook, but for a full delta sync.
        await syncService.runFullSyncForAllUsers(); // Will be implemented in sync.js
        res.status(202).send('Full sync initiated.');
    } catch (error) {
        console.error('Error initiating full sync:', error);
        res.status(500).send('Internal Server Error');
    }
});

// Endpoint for renewing Microsoft Graph subscriptions (Cloud Scheduler)
app.post('/graph/subscriptions/renew', async (req, res) => {
    try {
        console.log('Attempting to renew Microsoft Graph subscriptions.');
        // This endpoint will retrieve existing subscriptions and renew them.
        await syncService.renewGraphSubscriptions(); // Will be implemented in sync.js or graph.js
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

module.exports = app;
