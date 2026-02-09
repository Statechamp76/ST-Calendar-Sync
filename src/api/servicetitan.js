const { getSecrets } = require('../utils/secrets');
const { Interval, DateTime } = require('luxon');

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

    const baseUrl = `https://api.servicetitan.io/v2/tenant/${tenantId}`; // Using v2 API

    const defaultHeaders = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'ST-Tenant': tenantId,
        'Accept': 'application/json'
    };

    const config = {
        ...options,
        headers: {
            ...defaultHeaders,
            ...options.headers,
        },
    };

    try {
        const response = await fetch(`${baseUrl}${endpoint}`, config);

        if (!response.ok) {
            const errorBody = await response.text();
            throw new Error(`ServiceTitan API Error: ${response.status} ${response.statusText} - ${errorBody}`);
        }

        return response.json();
    } catch (error) {
        console.error(`Error calling ServiceTitan API endpoint ${endpoint}:`, error);
        throw error;
    }
}

/**
 * Creates a ServiceTitan Non-Job Appointment.
 * @param {object} appointmentData - Data for the new appointment.
 * @returns {Promise<string>} The ID of the newly created appointment.
 */
async function createNonJob(appointmentData) {
    console.log('Creating Non-Job Appointment:', appointmentData);
    // appointmentData should contain: technicianId, timesheetCodeId, start (ISO), duration (HH:mm:ss), name
    const payload = {
        technicianId: parseInt(appointmentData.technicianId),
        timesheetCodeId: parseInt(appointmentData.timesheetCodeId),
        start: appointmentData.start, // ISO 8601 string
        duration: appointmentData.duration, // HH:mm:ss
        name: appointmentData.name,
    };

    const response = await stApiRequest('/nonjobappointments', {
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
    console.log(`Updating Non-Job Appointment ${appointmentId}:`, updateData);
    const payload = {
        technicianId: parseInt(updateData.technicianId),
        timesheetCodeId: parseInt(updateData.timesheetCodeId),
        start: updateData.start,
        duration: updateData.duration,
        name: updateData.name,
    };

    await stApiRequest(`/nonjobappointments/${appointmentId}`, {
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
    await stApiRequest(`/nonjobappointments/${appointmentId}`, {
        method: 'DELETE',
    });
    console.log(`Non-Job Appointment ${appointmentId} deleted.`);
}

module.exports = {
    createNonJob,
    updateNonJob,
    deleteNonJob,
};
