const { getSecrets } = require('../utils/secrets');
const { DateTime } = require('luxon');

let accessTokenCache = {
    token: null,
    expiry: null,
};

/**
 * Retrieves a valid ServiceTitan access token, refreshing it if necessary.
 * Uses client credentials flow.
 * @returns {Promise<string>} The access token.
 */
async function getAccessToken() {
    // Check if token is still valid (give a 60-second buffer for network latency)
    if (accessTokenCache.token && accessTokenCache.expiry && DateTime.now().plus({ seconds: 60 }) < accessTokenCache.expiry) {
        return accessTokenCache.token;
    }

    console.log('ServiceTitan access token expired or not present. Refreshing...');
    
    const secrets = await getSecrets([
        'SERVICETITAN_CLIENT_ID',
        'SERVICETITAN_CLIENT_SECRET',
    ]);

    const clientId = secrets.SERVICETITAN_CLIENT_ID;
    const clientSecret = secrets.SERVICETITAN_CLIENT_SECRET;

    const tokenUrl = 'https://auth.servicetitan.io/connect/token';
    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);

    try {
        const response = await fetch(tokenUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: params.toString(),
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`ServiceTitan token refresh failed: ${response.status} - ${errorText}`);
        }

        const data = await response.json();
        accessTokenCache.token = data.access_token;
        // Calculate expiry based on expires_in (seconds)
        accessTokenCache.expiry = DateTime.now().plus({ seconds: data.expires_in });
        console.log('ServiceTitan access token refreshed successfully.');
        return accessTokenCache.token;

    } catch (error) {
        console.error('Error refreshing ServiceTitan token:', error);
        throw error;
    }
}

/**
 * Makes an authenticated request to the ServiceTitan API.
 * @param {string} endpoint - The API endpoint relative to the base URL.
 * @param {object} options - Fetch options.
 * @returns {Promise<object>} JSON response from the API.
 */
async function stApiRequest(endpoint, options = {}) {
    const token = await getAccessToken();
    const secrets = await getSecrets(['SERVICETITAN_TENANT_ID']);
    const tenantId = secrets.SERVICETITAN_TENANT_ID;
    const appKey = (process.env.SERVICETITAN_APP_KEY || '').trim();

    const baseUrl = `https://api.servicetitan.io/dispatch/v2/tenant/${tenantId}`;

    const defaultHeaders = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'ST-Tenant': tenantId,
        'Accept': 'application/json'
    };
    if (appKey) {
        defaultHeaders['ST-App-Key'] = appKey;
    }

    const config = {
        ...options,
        headers: {
            ...defaultHeaders,
            ...options.headers,
        },
    };

    let lastError;
    const maxAttempts = 4;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            const response = await fetch(`${baseUrl}${endpoint}`, config);

            if (response.status === 429 || response.status >= 500) {
                const retryBody = await response.text();
                lastError = new Error(`ServiceTitan API retryable error ${response.status}: ${retryBody}`);
                if (attempt < maxAttempts) {
                    await wait(getBackoffMs(attempt));
                    continue;
                }
                throw lastError;
            }

            if (!response.ok) {
                const errorBody = await response.text();
                const error = new Error(`ServiceTitan API Error: ${response.status} ${response.statusText} - ${errorBody}`);
                error.statusCode = response.status;
                throw error;
            }

            if (response.status === 204) {
                return null;
            }

            return response.json();
        } catch (error) {
            lastError = error;
            const isNetworkError = !Object.prototype.hasOwnProperty.call(error, 'statusCode');
            if (isNetworkError && attempt < maxAttempts) {
                await wait(getBackoffMs(attempt));
                continue;
            }
            throw error;
        }
    }

    console.error(`Error calling ServiceTitan API endpoint ${endpoint}:`, lastError);
    throw lastError;
}

function getBackoffMs(attempt) {
    const jitter = Math.floor(Math.random() * 100);
    return Math.min(5000, 250 * (2 ** (attempt - 1)) + jitter);
}

function wait(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Creates a ServiceTitan Non-Job Appointment.
 * @param {object} appointmentData - Data for the new appointment.
 * @returns {Promise<string>} The ID of the newly created appointment.
 */
async function createNonJob(appointmentData) {
    // Avoid logging appointment names (could contain customer/private data). Keep logs minimal.
    console.log('Creating Non-Job Appointment', {
        technicianId: appointmentData.technicianId,
        start: appointmentData.start,
        duration: appointmentData.duration,
        allDay: Boolean(appointmentData.allDay),
        showOnTechnicianSchedule: Boolean(appointmentData.showOnTechnicianSchedule),
    });
    // appointmentData should contain: technicianId, timesheetCodeId, start (ISO), duration (HH:mm:ss), name
    const payload = {
        technicianId: parseInt(appointmentData.technicianId),
        start: appointmentData.start, // ISO 8601 string
        duration: appointmentData.duration, // HH:mm:ss
        name: appointmentData.name,
        allDay: Boolean(appointmentData.allDay),
        showOnTechnicianSchedule: Boolean(appointmentData.showOnTechnicianSchedule),
        clearDispatchBoard: Boolean(appointmentData.clearDispatchBoard),
        clearTechnicianView: Boolean(appointmentData.clearTechnicianView),
        removeTechnicianFromCapacityPlanning: Boolean(appointmentData.removeTechnicianFromCapacityPlanning),
        active: appointmentData.active !== false,
    };

    const timesheetCodeId = Number.parseInt(String(appointmentData.timesheetCodeId || ''), 10);
    if (Number.isFinite(timesheetCodeId) && timesheetCodeId > 0) {
        payload.timesheetCodeId = timesheetCodeId;
    }

    const response = await stApiRequest('/non-job-appointments', {
        method: 'POST',
        body: JSON.stringify(payload),
    });
    console.log('Non-Job Appointment created:', response.id);
    return response.id;
}

/**
 * Updates an existing ServiceTitan Non-Job Appointment.
 * @param {string} appointmentId - The ID of the appointment to update.
 * @param {object} updateData - Data to update the appointment with.
 * @returns {Promise<void>}
 */
async function updateNonJob(appointmentId, updateData) {
    console.log(`Updating Non-Job Appointment ${appointmentId}`, {
        technicianId: updateData.technicianId,
        start: updateData.start,
        duration: updateData.duration,
        allDay: Boolean(updateData.allDay),
        showOnTechnicianSchedule: Boolean(updateData.showOnTechnicianSchedule),
    });
    const payload = {
        technicianId: parseInt(updateData.technicianId),
        start: updateData.start,
        duration: updateData.duration,
        name: updateData.name,
        allDay: Boolean(updateData.allDay),
        showOnTechnicianSchedule: Boolean(updateData.showOnTechnicianSchedule),
        clearDispatchBoard: Boolean(updateData.clearDispatchBoard),
        clearTechnicianView: Boolean(updateData.clearTechnicianView),
        removeTechnicianFromCapacityPlanning: Boolean(updateData.removeTechnicianFromCapacityPlanning),
        active: updateData.active !== false,
    };

    const timesheetCodeId = Number.parseInt(String(updateData.timesheetCodeId || ''), 10);
    if (Number.isFinite(timesheetCodeId) && timesheetCodeId > 0) {
        payload.timesheetCodeId = timesheetCodeId;
    }

    await stApiRequest(`/non-job-appointments/${appointmentId}`, {
        method: 'PUT',
        body: JSON.stringify(payload),
    });
    console.log(`Non-Job Appointment ${appointmentId} updated.`);
}

/**
 * Deletes a ServiceTitan Non-Job Appointment.
 * @param {string} appointmentId - The ID of the appointment to delete.
 * @returns {Promise<void>}
 */
async function deleteNonJob(appointmentId) {
    console.log(`Deleting Non-Job Appointment ${appointmentId}`);
    try {
        await stApiRequest(`/non-job-appointments/${appointmentId}`, {
            method: 'DELETE',
        });
        console.log(`Non-Job Appointment ${appointmentId} deleted.`);
    } catch (error) {
        if (error && error.statusCode === 404) {
            console.warn(`Non-Job Appointment ${appointmentId} already missing (404).`);
            return;
        }
        throw error;
    }
}

module.exports = {
    createNonJob,
    updateNonJob,
    deleteNonJob,
};
