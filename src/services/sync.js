const graph = require('../api/graph');
const servicetitan = require('../api/servicetitan');
const sheets = require('./sheets');
const { splitMultiDayEvent, TIMEZONE } = require('../utils/time');
const { getSecrets } = require('../utils/secrets');
const crypto = require('crypto');
const { DateTime, Duration, Interval } = require('luxon');

/**
 * Generates a hash for an Outlook event based on its significant properties.
 * This hash is used to detect changes and ensure idempotency.
 * @param {object} event - The Outlook event object.
 * @returns {string} MD5 hash of the event's relevant data.
 */
function generateEventHash(event) {
    const data = {
        subject: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        showAs: event.showAs,
        isPrivate: event.isPrivate,
        location: event.location ? event.location.displayName : '',
        body: event.body ? event.body.content : '',
        // Add other fields that, if changed, should trigger a ServiceTitan update
    };
    return crypto.createHash('md5').update(JSON.stringify(data)).digest('hex');
}

/**
 * Runs a delta synchronization for a single user's mailbox.
 * This is the core worker function called by the Pub/Sub subscriber.
 * @param {string} userUpn - The User Principal Name (email) of the user to sync.
 */
async function runDeltaSyncForUser(userUpn) {
    console.log(`Starting delta sync for user: ${userUpn}`);

    // 1. Get user configuration from TechMap
    const techMap = await sheets.getTechMap();
    const userConfig = techMap.find(u => u.outlook_upn === userUpn);

    if (!userConfig || !userConfig.enabled) {
        console.log(`User ${userUpn} not found or not enabled in TechMap. Skipping sync.`);
        return;
    }

    // 2. Get deltaLink for this user
    let deltaState = await sheets.getDeltaState(userUpn);
    const initialDeltaLink = deltaState.delta_link;
    console.log(`Using deltaLink for ${userUpn}: ${initialDeltaLink || 'None (initial sync)'}`);

    // 3. Fetch changes from Microsoft Graph
    let graphResponse;
    try {
        graphResponse = await graph.getDeltaEvents(userUpn, initialDeltaLink);
    } catch (error) {
        console.error(`Error fetching delta events for ${userUpn}:`, error);
        // If initial sync fails due to bad delta link, try a full initial sync (without delta link)
        if (initialDeltaLink && error.message.includes('DeltaTokenNotFind') || error.message.includes('ResyncRequired')) {
            console.warn(`Delta link invalid for ${userUpn}. Attempting full initial sync.`);
            graphResponse = await graph.getDeltaEvents(userUpn, null);
        } else {
            throw error; // Re-throw if it's a different error
        }
    }
    
    const { events, nextDeltaLink } = graphResponse;
    console.log(`Received ${events.length} events from Graph for ${userUpn}.`);

    for (const event of events) {
        const outlookEventId = event.id;
        const existingMapping = await sheets.findEventMapping(userUpn, outlookEventId);
        
        if (event['@removed']) {
            // --- Handle Deletion ---
            console.log(`Outlook event ${outlookEventId} for ${userUpn} was removed.`);
            if (existingMapping) {
                const stNonJobIds = JSON.parse(existingMapping.st_nonjob_ids_json || '[]');
                for (const stId of stNonJobIds) {
                    try {
                        await servicetitan.deleteNonJob(stId);
                        console.log(`Deleted ServiceTitan non-job ${stId} for event ${outlookEventId}.`);
                    } catch (delError) {
                        console.error(`Failed to delete ST non-job ${stId} for ${outlookEventId}:`, delError);
                    }
                }
                await sheets.deleteEventMapping(userUpn, outlookEventId); // Mark as deleted in sheets
            } else {
                console.log(`No existing mapping for removed event ${outlookEventId}. Skipping delete.`);
            }
            continue; // Move to next event
        }

        // --- Filter Events for Sync ---
        // Ignore events where showAs = free. If it was previously synced, it needs to be deleted.
        if (event.showAs === 'free') {
             console.log(`Outlook event ${outlookEventId} for ${userUpn} is 'free'.`);
             if (existingMapping) {
                 console.log(`Event ${outlookEventId} was previously synced. Deleting ST non-jobs.`);
                 const stNonJobIds = JSON.parse(existingMapping.st_nonjob_ids_json || '[]');
                 for (const stId of stNonJobIds) {
                     try {
                         await servicetitan.deleteNonJob(stId);
                     } catch (delError) {
                         console.error(`Failed to delete ST non-job ${stId} for event ${outlookEventId}:`, delError);
                     }
                 }
                 await sheets.deleteEventMapping(userUpn, outlookEventId);
             }
             continue; // Move to next event
        }
        
        // --- Process Create/Update ---
        const currentEventHash = generateEventHash(event);
        if (existingMapping && existingMapping.last_hash === currentEventHash) {
            console.log(`Event ${outlookEventId} for ${userUpn} is unchanged (hash match). Skipping update.`);
            continue;
        }

        console.log(`Processing update/create for Outlook event ${outlookEventId} for ${userUpn}.`);

        // Determine ServiceTitan appointment name
        const stAppointmentName = event.isPrivate ? "Busy" : (event.subject || "Calendar Event");

        // Split multi-day events into single-day blocks
        const eventBlocks = splitMultiDayEvent(event.start.dateTime, event.end.dateTime);
        const newStNonJobIds = [];
        const oldStNonJobIds = existingMapping ? JSON.parse(existingMapping.st_nonjob_ids_json || '[]') : [];

        for (let i = 0; i < eventBlocks.length; i++) {
            const block = eventBlocks[i];
            const startDateTime = DateTime.fromISO(block.start, { zone: TIMEZONE });
            const endDateTime = DateTime.fromISO(block.end, { zone: TIMEZONE });
            const duration = Interval.fromDateTimes(startDateTime, endDateTime).toDuration().toFormat('hh:mm:ss');
            
            const stAppointmentData = {
                technicianId: userConfig.st_technician_id,
                timesheetCodeId: userConfig.st_timesheet_code_id,
                start: startDateTime.toISO(),
                duration: duration,
                name: stAppointmentName,
            };

            let currentStId = oldStNonJobIds[i]; // Try to reuse existing ID

            if (currentStId) {
                try {
                    await servicetitan.updateNonJob(currentStId, stAppointmentData);
                    console.log(`Updated ST non-job ${currentStId} for event ${outlookEventId}.`);
                } catch (updateError) {
                    console.error(`Failed to update ST non-job ${currentStId} for ${outlookEventId}:`, updateError);
                    // If update fails, perhaps the ST non-job was deleted externally, try creating a new one.
                    currentStId = await servicetitan.createNonJob(stAppointmentData);
                    console.log(`Created new ST non-job ${currentStId} after failed update for ${outlookEventId}.`);
                }
            } else {
                currentStId = await servicetitan.createNonJob(stAppointmentData);
                console.log(`Created new ST non-job ${currentStId} for event ${outlookEventId}.`);
            }
            newStNonJobIds.push(currentStId);
        }

        // Delete any remaining old ST non-jobs that are no longer part of this event (e.g., event shortened)
        for (let i = eventBlocks.length; i < oldStNonJobIds.length; i++) {
            const stIdToDelete = oldStNonJobIds[i];
            try {
                await servicetitan.deleteNonJob(stIdToDelete);
                console.log(`Deleted surplus ST non-job ${stIdToDelete} for event ${outlookEventId}.`);
            } catch (delError) {
                console.error(`Failed to delete surplus ST non-job ${stIdToDelete} for ${outlookEventId}:`, delError);
            }
        }

        // 4. Update EventMap
        await sheets.updateEventMapping(userUpn, outlookEventId, newStNonJobIds, currentEventHash, 'SYNCED', existingMapping ? existingMapping.rowIndex : null);
    }

    // 5. Update DeltaState with the new deltaLink
    await sheets.updateDeltaState(userUpn, nextDeltaLink, deltaState.rowIndex);
    console.log(`Delta sync completed for user: ${userUpn}. Next deltaLink: ${nextDeltaLink}`);
}

/**
 * Initiates a full synchronization for all enabled users in TechMap.
 * This is primarily for the nightly backstop Cloud Scheduler job.
 */
async function runFullSyncForAllUsers() {
    console.log('Starting full sync for all enabled users.');
    const techMap = await sheets.getTechMap();
    const pubsub = new PubSub();
    const topicName = 'graph-notifications';

    for (const userConfig of techMap) {
        if (userConfig.enabled) {
            console.log(`Initiating sync for ${userConfig.outlook_upn} via Pub/Sub.`);
            // Publish message to trigger individual user sync
            await pubsub.topic(topicName).publishMessage({ json: { upn: userConfig.outlook_upn } });
        }
    }
    console.log('Full sync initiation complete. Individual user syncs will proceed via Pub/Sub.');
}


/**
 * Renews all active Microsoft Graph subscriptions.
 * This is called by a Cloud Scheduler job.
 */
async function renewGraphSubscriptions() {
    console.log('Starting Graph subscription renewal process.');
    const techMap = await sheets.getTechMap();
    const secrets = await getSecrets(['GRAPH_WEBHOOK_URL', 'GRAPH_CLIENT_STATE']);
    const notificationUrl = secrets.GRAPH_WEBHOOK_URL;
    const clientState = secrets.GRAPH_CLIENT_STATE;

    if (!notificationUrl) {
        console.error('GRAPH_WEBHOOK_URL secret is not configured. Cannot renew subscriptions.');
        return;
    }

    for (const userConfig of techMap) {
        if (userConfig.enabled) {
            try {
                // The createOrRenewSubscription function handles checking for existing subscriptions
                await graph.createOrRenewSubscription(userConfig.outlook_upn, notificationUrl, clientState);
                console.log(`Subscription renewed for ${userConfig.outlook_upn}.`);
            } catch (error) {
                console.error(`Failed to renew subscription for ${userConfig.outlook_upn}:`, error);
            }
        }
    }
    console.log('Graph subscription renewal process completed.');
}


module.exports = {
    runDeltaSyncForUser,
    runFullSyncForAllUsers,
    renewGraphSubscriptions,
};
