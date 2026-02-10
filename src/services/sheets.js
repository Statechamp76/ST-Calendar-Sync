const { google } = require('googleapis');
const { getSecrets } = require('../utils/secrets');
const { DateTime } = require('luxon');

let sheetsService;
let spreadsheetId;

// Simple in-memory cache to reduce Sheets read quota pressure during large sync runs.
// A longer TTL is intentional; this service is the sole writer for these sheets in normal operation.
const CACHE_TTL_MS = 5 * 60_000;
let eventMapCache = {
    loadedAtMs: 0,
    headerRow: null,
    rows: null,
};

// --- Initialization ---
async function initializeSheets() {
    if (sheetsService) return; // Already initialized

    const secrets = await getSecrets(['GOOGLE_SPREADSHEET_ID']);
    spreadsheetId = secrets.GOOGLE_SPREADSHEET_ID;

    // Authenticate using Google Cloud's default credentials.
    // This will pick up credentials from the Cloud Run service account.
    const auth = new google.auth.GoogleAuth({
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const authClient = await auth.getClient();
    sheetsService = google.sheets({ version: 'v4', auth: authClient });
    console.log('Google Sheets API client initialized.');
}

function invalidateEventMapCache() {
    eventMapCache.loadedAtMs = 0;
    eventMapCache.headerRow = null;
    eventMapCache.rows = null;
}

// --- Generic Sheet Read/Write Helpers ---

/**
 * Reads rows from a specified sheet range.
 * @param {string} range - The A1 notation or R1C1 notation of the range to retrieve.
 * @returns {Promise<Array<Array<string>>>} A 2D array of values from the sheet.
 */
async function readSheetRows(range) {
    await initializeSheets();
    try {
        const response = await sheetsService.spreadsheets.values.get({
            spreadsheetId,
            range,
        });
        return response.data.values || [];
    } catch (error) {
        console.error(`Error reading from sheet range "${range}":`, error.message);
        throw new Error(`Failed to read from Google Sheet: ${range}`);
    }
}

/**
 * Appends a row to a specified sheet.
 * @param {string} range - The A1 notation of the sheet to append to (e.g., 'Sheet1!A1').
 * @param {Array<string>} rowData - An array of values for the new row.
 * @returns {Promise<void>}
 */
async function appendSheetRow(range, rowData) {
    await initializeSheets();
    try {
        await sheetsService.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: [rowData],
            },
        });
    } catch (error) {
        console.error(`Error appending to sheet range "${range}":`, error.message);
        throw new Error(`Failed to append row to Google Sheet: ${range}`);
    }
}

/**
 * Updates a specific cell or range in a sheet.
 * @param {string} range - The A1 notation of the cell or range to update.
 * @param {Array<Array<string>>} values - The new values to write.
 * @returns {Promise<void>}
 */
async function updateSheetRange(range, values) {
    await initializeSheets();
    try {
        await sheetsService.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: values,
            },
        });
    } catch (error) {
        console.error(`Error updating sheet range "${range}":`, error.message);
        throw new Error(`Failed to update Google Sheet: ${range}`);
    }
}

/**
 * Deletes rows from a specified sheet.
 * @param {string} sheetName - The name of the sheet (e.g., 'EventMap').
 * @param {number} startRowIndex - The 1-based index of the first row to delete.
 * @param {number} endRowIndex - The 1-based index of the last row to delete (exclusive).
 * @returns {Promise<void>}
 */
async function deleteSheetRows(sheetName, startRowIndex, endRowIndex) {
    await initializeSheets();
    try {
        // Need to get the sheetId first
        const metadata = await sheetsService.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties',
        });
        const sheet = metadata.data.sheets.find(s => s.properties.title === sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found.`);
        }
        const sheetId = sheet.properties.sheetId;

        await sheetsService.spreadsheets.batchUpdate({
            spreadsheetId,
            resource: {
                requests: [{
                    deleteDimension: {
                        range: {
                            sheetId: sheetId,
                            dimension: 'ROWS',
                            startIndex: startRowIndex - 1, // API is 0-indexed
                            endIndex: endRowIndex - 1,
                        },
                    },
                }],
            },
        });
    } catch (error) {
        console.error(`Error deleting rows from sheet "${sheetName}" (rows ${startRowIndex}-${endRowIndex}):`, error.message);
        throw new Error(`Failed to delete rows from Google Sheet: ${sheetName}`);
    }
}

// --- Specific Data Access Helpers ---

/**
 * Retrieves the TechMap configuration.
 * @returns {Promise<Array<object>>} Array of technician mappings.
 */
async function getTechMap() {
    const rows = await readSheetRows('TechMap!A2:D'); // Assuming headers are in A1:D1
    return rows.map(row => ({
        outlook_upn: row[0] || '',
        st_technician_id: row[1] || '',
        st_timesheet_code_id: row[2] || '',
        enabled: (row[3] || 'FALSE').toUpperCase() === 'TRUE',
    }));
}

function getRequiredHeaderIndex(headerRowValues, headerName, sheetName) {
    if (!Array.isArray(headerRowValues) || headerRowValues.length === 0) {
        throw new Error(`Missing header row in ${sheetName}.`);
    }

    const index = headerRowValues.indexOf(headerName);
    if (index === -1) {
        throw new Error(`Missing required header "${headerName}" in ${sheetName}.`);
    }

    return index;
}

/**
 * Retrieves delta state for a specific UPN.
 * @param {string} outlookUpn - The UPN to retrieve delta state for.
 * @returns {Promise<object>} Delta state object, or a default if not found.
 */
async function getDeltaState(outlookUpn) {
    const rows = await readSheetRows('DeltaState!A2:E'); // Assuming headers in A1:E1
    const headerRowValues = (await readSheetRows('DeltaState!A1:E1'))[0];
    const upnIndex = getRequiredHeaderIndex(headerRowValues, 'outlook_upn', 'DeltaState');

    let rowIndex = -1;
    const existingEntry = rows.find((row, idx) => {
        if (row[upnIndex] === outlookUpn) {
            rowIndex = idx + 2; // +2 because header is 1 and array is 0-indexed
            return true;
        }
        return false;
    });

    if (existingEntry) {
        return {
            rowIndex: rowIndex,
            outlook_upn: existingEntry[0],
            delta_link: existingEntry[1],
            window_end: existingEntry[2],
            last_run_utc: existingEntry[3],
        };
    } else {
        // Return default for initial sync
        return {
            rowIndex: null, // Indicates new entry
            outlook_upn: outlookUpn,
            delta_link: null,
            window_end: null, // Or DateTime.now().plus({days: 90}).toISO() depending on initial window
            last_run_utc: null,
        };
    }
}

/**
 * Updates or creates delta state for a specific UPN.
 * @param {string} outlookUpn - The UPN to update delta state for.
 * @param {string} newDeltaLink - The new delta link.
 * @param {number | null} existingRowIndex - The 1-based row index if updating an existing entry.
 * @returns {Promise<void>}
 */
async function updateDeltaState(outlookUpn, newDeltaLink, existingRowIndex) {
    const now = DateTime.utc().toISO();
    const rowData = [
        outlookUpn,
        newDeltaLink,
        // Assuming window_end is managed by the Graph API or a separate logic
        '', // Placeholder for window_end
        now, // last_run_utc
    ];

    if (existingRowIndex) {
        // Update existing row (e.g., DeltaState!A<rowIndex>:D<rowIndex>)
        await updateSheetRange(`DeltaState!A${existingRowIndex}`, [rowData]);
    } else {
        // Append new row
        await appendSheetRow('DeltaState!A:D', rowData);
    }
    console.log(`Delta state updated for ${outlookUpn}.`);
}


/**
 * Finds an event mapping entry.
 * @param {string} outlookUpn - The UPN.
 * @param {string} outlookEventId - The Outlook event ID.
 * @returns {Promise<object | null>} The event mapping object with its row index, or null if not found.
 */
async function findEventMapping(outlookUpn, outlookEventId) {
    const nowMs = Date.now();
    if (!eventMapCache.rows || nowMs - eventMapCache.loadedAtMs > CACHE_TTL_MS) {
        const [rows, header] = await Promise.all([
            readSheetRows('EventMap!A2:F'),
            readSheetRows('EventMap!A1:F1'),
        ]);
        eventMapCache.rows = rows;
        eventMapCache.headerRow = header[0];
        eventMapCache.loadedAtMs = nowMs;
    }

    const rows = eventMapCache.rows;
    const headerRowValues = eventMapCache.headerRow;
    const upnIndex = getRequiredHeaderIndex(headerRowValues, 'outlook_upn', 'EventMap');
    const eventIdIndex = getRequiredHeaderIndex(headerRowValues, 'outlook_event_id', 'EventMap');

    let rowIndex = -1;
    const existingEntry = rows.find((row, idx) => {
        if (row[upnIndex] === outlookUpn && row[eventIdIndex] === outlookEventId) {
            rowIndex = idx + 2; // +2 for header row and 0-index adjustment
            return true;
        }
        return false;
    });

    if (existingEntry) {
        return {
            rowIndex: rowIndex,
            outlook_upn: existingEntry[0],
            outlook_event_id: existingEntry[1],
            st_nonjob_ids_json: existingEntry[2],
            last_hash: existingEntry[3],
            last_synced_utc: existingEntry[4],
            status: existingEntry[5],
        };
    }
    return null;
}

/**
 * Updates or creates an event mapping entry.
 * @param {string} outlookUpn - The UPN.
 * @param {string} outlookEventId - The Outlook event ID.
 * @param {Array<string>} stNonJobIds - Array of ServiceTitan non-job appointment IDs.
 * @param {string} lastHash - MD5 hash of the event content.
 * @param {string} status - Current status of the event (e.g., 'SYNCED', 'DELETED').
 * @param {number | null} existingRowIndex - The 1-based row index if updating.
 * @returns {Promise<void>}
 */
async function updateEventMapping(outlookUpn, outlookEventId, stNonJobIds, lastHash, status = 'SYNCED', existingRowIndex) {
    const now = DateTime.utc().toISO();
    const rowData = [
        outlookUpn,
        outlookEventId,
        JSON.stringify(stNonJobIds), // Store as JSON string
        lastHash,
        now,
        status,
    ];

    if (existingRowIndex) {
        await updateSheetRange(`EventMap!A${existingRowIndex}`, [rowData]);
    } else {
        await appendSheetRow('EventMap!A:F', rowData);
    }

    // Keep cache warm to avoid read-quota bursts during backfills.
    if (eventMapCache.rows) {
        if (existingRowIndex) {
            const idx = existingRowIndex - 2;
            if (idx >= 0 && idx < eventMapCache.rows.length) {
                eventMapCache.rows[idx] = rowData;
            }
        } else {
            eventMapCache.rows.push(rowData);
        }
        eventMapCache.loadedAtMs = Date.now();
    }
    console.log(`Event mapping updated for ${outlookUpn}:${outlookEventId}.`);
}

/**
 * Deletes an event mapping entry by marking it as DELETED or physically removing it.
 * For production, marking as 'DELETED' is safer for auditing.
 * @param {string} outlookUpn - The UPN.
 * @param {string} outlookEventId - The Outlook event ID.
 * @returns {Promise<void>}
 */
async function deleteEventMapping(outlookUpn, outlookEventId) {
    const existingMapping = arguments.length >= 3 ? arguments[2] : await findEventMapping(outlookUpn, outlookEventId);
    if (existingMapping && existingMapping.rowIndex) {
        // Option 1: Mark as DELETED (recommended for auditing)
        await updateEventMapping(
            outlookUpn,
            outlookEventId,
            [], // Clear ST IDs
            existingMapping.last_hash, // Keep hash or update to a 'deleted' hash
            'DELETED',
            existingMapping.rowIndex
        );
        console.log(`Event mapping for ${outlookUpn}:${outlookEventId} marked as DELETED.`);
        
        // Option 2: Physically delete the row (use with caution)
        // await deleteSheetRows('EventMap', existingMapping.rowIndex, existingMapping.rowIndex + 1);
        // console.log(`Event mapping for ${outlookUpn}:${outlookEventId} physically deleted.`);

    } else {
        console.warn(`Attempted to delete non-existent event mapping for ${outlookUpn}:${outlookEventId}.`);
    }
}


module.exports = {
    getTechMap,
    getDeltaState,
    updateDeltaState,
    findEventMapping,
    updateEventMapping,
    deleteEventMapping,
    readSheetRows, // Exposed for runFullSyncForAllUsers might need it
};
