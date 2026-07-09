const { google } = require('googleapis');
const { getSecrets } = require('../utils/secrets');
const { DateTime } = require('luxon');
const { normalizeUpn } = require('../utils/upn');

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
let techMapCache = {
    loadedAtMs: 0,
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

async function getSheetMetadata() {
    await initializeSheets();
    return withRetry(async () => {
        return sheetsService.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties',
        });
    }, 'get metadata');
}

async function ensureSheetExists(sheetName, headerRowValues) {
    await initializeSheets();
    const metadata = await getSheetMetadata();
    const existing = (metadata.data.sheets || []).find((s) => s.properties && s.properties.title === sheetName);
    if (existing) return;

    await withRetry(async () => {
        await sheetsService.spreadsheets.batchUpdate({
            spreadsheetId,
            resource: {
                requests: [{
                    addSheet: {
                        properties: {
                            title: sheetName,
                        },
                    },
                }],
            },
        });
    }, `add sheet ${sheetName}`);

    if (Array.isArray(headerRowValues) && headerRowValues.length > 0) {
        await updateSheetRange(`${sheetName}!A1`, [headerRowValues]);
    }
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
async function readSheetRows(range, spreadsheetIdOverride = null) {
    await initializeSheets();
    const targetSpreadsheetId = spreadsheetIdOverride || spreadsheetId;
    return withRetry(async () => {
        const response = await sheetsService.spreadsheets.values.get({
            spreadsheetId: targetSpreadsheetId,
            range,
        });
        return response.data.values || [];
    }, `read ${range}`);
}

/**
 * Appends a row to a specified sheet.
 * @param {string} range - The A1 notation of the sheet to append to (e.g., 'Sheet1!A1').
 * @param {Array<string>} rowData - An array of values for the new row.
 * @returns {Promise<void>}
 */
async function appendSheetRow(range, rowData) {
    await initializeSheets();
    return withRetry(async () => {
        await sheetsService.spreadsheets.values.append({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: [rowData],
            },
        });
    }, `append ${range}`);
}

async function appendSheetRows(range, rowsData, spreadsheetIdOverride = null) {
    const rows = Array.isArray(rowsData) ? rowsData : [];
    if (rows.length === 0) return;
    await initializeSheets();
    const targetSpreadsheetId = spreadsheetIdOverride || spreadsheetId;
    return withRetry(async () => {
        await sheetsService.spreadsheets.values.append({
            spreadsheetId: targetSpreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: rows,
            },
        });
    }, `append ${range} (batch ${rows.length})`);
}

/**
 * Updates a specific cell or range in a sheet.
 * @param {string} range - The A1 notation of the cell or range to update.
 * @param {Array<Array<string>>} values - The new values to write.
 * @returns {Promise<void>}
 */
async function updateSheetRange(range, values) {
    await initializeSheets();
    return withRetry(async () => {
        await sheetsService.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: values,
            },
        });
    }, `update ${range}`);
}

async function batchUpdateSheetRanges(data, spreadsheetIdOverride = null) {
    const updates = Array.isArray(data) ? data : [];
    if (updates.length === 0) return;
    await initializeSheets();
    const targetSpreadsheetId = spreadsheetIdOverride || spreadsheetId;
    return withRetry(async () => {
        await sheetsService.spreadsheets.values.batchUpdate({
            spreadsheetId: targetSpreadsheetId,
            resource: {
                valueInputOption: 'RAW',
                data: updates,
            },
        });
    }, `batchUpdate values (${updates.length} ranges)`);
}

async function clearSheetRange(range) {
    await initializeSheets();
    return withRetry(async () => {
        await sheetsService.spreadsheets.values.clear({
            spreadsheetId,
            range,
        });
    }, `clear ${range}`);
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
    return withRetry(async () => {
        // Need to get the sheetId first
        const metadata = await getSheetMetadata();
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
    }, `delete rows ${sheetName} ${startRowIndex}-${endRowIndex}`);
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

function getRetryableCode(error) {
    // googleapis errors often include `code` and a structured response.
    if (error && typeof error.code === 'number') return error.code;
    const status = error?.response?.status;
    if (typeof status === 'number') return status;
    return null;
}

function isRetryableSheetsError(error) {
    const code = getRetryableCode(error);
    if (code === 429) return true;
    if (code >= 500 && code <= 599) return true;
    const msg = String(error?.message || '').toLowerCase();
    if (msg.includes('quota')) return true;
    if (msg.includes('rate limit')) return true;
    if (msg.includes('backend error')) return true;
    return false;
}

async function withRetry(fn, label) {
    const maxAttempts = 6;
    let lastErr = null;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        try {
            return await fn();
        } catch (error) {
            lastErr = error;
            const retryable = isRetryableSheetsError(error);
            const code = getRetryableCode(error);
            if (!retryable || attempt === maxAttempts) {
                const details = code ? `code=${code}` : 'code=unknown';
                throw new Error(`Google Sheets ${label} failed (${details}): ${error.message}`);
            }

            // Exponential backoff with jitter.
            const base = Math.min(10_000, 500 * (2 ** (attempt - 1)));
            const jitter = Math.floor(Math.random() * 250);
            const waitMs = base + jitter;
            console.warn('sheets.retry', { label, attempt, waitMs, code });
            await sleep(waitMs);
        }
    }

    throw lastErr;
}

// --- Specific Data Access Helpers ---

/**
 * Retrieves the TechMap configuration.
 * @returns {Promise<Array<object>>} Array of technician mappings.
 */
async function getTechMap() {
    const nowMs = Date.now();
    if (!techMapCache.rows || nowMs - techMapCache.loadedAtMs > CACHE_TTL_MS) {
        techMapCache.rows = await readSheetRows('TechMap!A2:D');
        techMapCache.loadedAtMs = nowMs;
    }
    return techMapCache.rows.map(row => ({
        outlook_upn: normalizeUpn(row[0]),
        st_technician_id: String(row[1] || '').trim(),
        st_timesheet_code_id: String(row[2] || '').trim(),
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
    const normalizedUpn = normalizeUpn(outlookUpn);
    const rows = await readSheetRows('DeltaState!A2:E'); // Assuming headers in A1:E1
    const headerRowValues = (await readSheetRows('DeltaState!A1:E1'))[0];
    const upnIndex = getRequiredHeaderIndex(headerRowValues, 'outlook_upn', 'DeltaState');

    let rowIndex = -1;
    const existingEntry = rows.find((row, idx) => {
        if (normalizeUpn(row[upnIndex]) === normalizedUpn) {
            rowIndex = idx + 2; // +2 because header is 1 and array is 0-indexed
            return true;
        }
        return false;
    });

    if (existingEntry) {
        return {
            rowIndex: rowIndex,
            outlook_upn: normalizeUpn(existingEntry[0]),
            delta_link: existingEntry[1],
            window_end: existingEntry[2],
            last_run_utc: existingEntry[3],
        };
    } else {
        // Return default for initial sync
        return {
            rowIndex: null, // Indicates new entry
            outlook_upn: normalizedUpn,
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
    const normalizedUpn = normalizeUpn(outlookUpn);
    const now = DateTime.utc().toISO();
    const rowData = [
        normalizedUpn,
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
    console.log(`Delta state updated for ${normalizedUpn}.`);
}


/**
 * Finds an event mapping entry.
 * @param {string} outlookUpn - The UPN.
 * @param {string} outlookEventId - The Outlook event ID.
 * @returns {Promise<object | null>} The event mapping object with its row index, or null if not found.
 */
async function findEventMapping(outlookUpn, outlookEventId) {
    const normalizedUpn = normalizeUpn(outlookUpn);
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
        if (normalizeUpn(row[upnIndex]) === normalizedUpn && row[eventIdIndex] === outlookEventId) {
            rowIndex = idx + 2; // +2 for header row and 0-index adjustment
            return true;
        }
        return false;
    });

    if (existingEntry) {
        return {
            rowIndex: rowIndex,
            outlook_upn: normalizeUpn(existingEntry[0]),
            outlook_event_id: existingEntry[1],
            st_nonjob_ids_json: existingEntry[2],
            last_hash: existingEntry[3],
            last_synced_utc: existingEntry[4],
            status: existingEntry[5],
        };
    }
    return null;
}

// --- Distributed Locks (Sheets-backed) ---
//
// Locks sheet schema:
// A: lock_key (e.g. outlook_upn)
// B: lease_until_utc (ISO)
// C: holder
// D: updated_utc (ISO)
const LOCKS_SHEET = 'Locks';
const LOCKS_HEADER = ['lock_key', 'lease_until_utc', 'holder', 'updated_utc'];

async function ensureLocksSheet() {
    await ensureSheetExists(LOCKS_SHEET, LOCKS_HEADER);
}

async function readLocksRows() {
    await ensureLocksSheet();
    const [rows, header] = await Promise.all([
        readSheetRows(`${LOCKS_SHEET}!A2:D`),
        readSheetRows(`${LOCKS_SHEET}!A1:D1`),
    ]);
    return { rows, header: (header[0] || []) };
}

async function findLockRow(lockKey) {
    const { rows, header } = await readLocksRows();
    const keyIndex = getRequiredHeaderIndex(header, 'lock_key', LOCKS_SHEET);

    let rowIndex = null;
    const row = rows.find((r, idx) => {
        if (String(r[keyIndex] || '') === String(lockKey || '')) {
            rowIndex = idx + 2;
            return true;
        }
        return false;
    });

    return { rowIndex, row: row || null, header };
}

function parseIsoToMs(value) {
    const s = String(value || '').trim();
    if (!s) return 0;
    const dt = DateTime.fromISO(s, { zone: 'utc' });
    if (!dt.isValid) return 0;
    return dt.toMillis();
}

async function tryAcquireLock(lockKey, holder, ttlSeconds) {
    const key = String(lockKey || '').trim();
    const h = String(holder || '').trim();
    const ttl = Number.parseInt(String(ttlSeconds || ''), 10);
    if (!key) throw new Error('Missing lockKey');
    if (!h) throw new Error('Missing holder');
    if (!Number.isFinite(ttl) || ttl <= 0) throw new Error('Invalid ttlSeconds');

    const now = DateTime.utc();
    const desiredLease = now.plus({ seconds: ttl }).toISO();
    const updated = now.toISO();

    const { rowIndex, row } = await findLockRow(key);
    const currentLeaseMs = row ? parseIsoToMs(row[1]) : 0;
    const nowMs = now.toMillis();

    // Only attempt takeover if expired (or missing).
    if (row && currentLeaseMs > nowMs) {
        return { acquired: false, reason: 'locked' };
    }

    const rowData = [key, desiredLease, h, updated];
    if (rowIndex) {
        await updateSheetRange(`${LOCKS_SHEET}!A${rowIndex}:D${rowIndex}`, [rowData]);
    } else {
        await appendSheetRow(`${LOCKS_SHEET}!A:D`, rowData);
    }

    // Verify we actually hold the lock (mitigates races).
    const verify = await findLockRow(key);
    const verifyHolder = verify.row ? String(verify.row[2] || '') : '';
    if (verifyHolder !== h) {
        return { acquired: false, reason: 'lost_race' };
    }
    return { acquired: true };
}

async function releaseLock(lockKey, holder) {
    const key = String(lockKey || '').trim();
    const h = String(holder || '').trim();
    if (!key) throw new Error('Missing lockKey');
    if (!h) throw new Error('Missing holder');

    const now = DateTime.utc().toISO();
    const { rowIndex, row } = await findLockRow(key);
    if (!rowIndex || !row) return { released: false, reason: 'missing' };

    const currentHolder = String(row[2] || '');
    if (currentHolder !== h) {
        return { released: false, reason: 'not_holder' };
    }

    const rowData = [key, now, h, now];
    await updateSheetRange(`${LOCKS_SHEET}!A${rowIndex}:D${rowIndex}`, [rowData]);
    return { released: true };
}

async function findEventMappingByGraphId(outlookUpn, graphEventId) {
    const normalizedUpn = normalizeUpn(outlookUpn);
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
    const statusIndex = getRequiredHeaderIndex(headerRowValues, 'status', 'EventMap');

    const needle = `gid=${graphEventId}`;
    let rowIndex = -1;
    const existingEntry = rows.find((row, idx) => {
        if (normalizeUpn(row[upnIndex]) === normalizedUpn && String(row[statusIndex] || '').includes(needle)) {
            rowIndex = idx + 2;
            return true;
        }
        return false;
    });

    if (existingEntry) {
        return {
            rowIndex: rowIndex,
            outlook_upn: normalizeUpn(existingEntry[0]),
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
    const normalizedUpn = normalizeUpn(outlookUpn);
    const now = DateTime.utc().toISO();
    const rowData = [
        normalizedUpn,
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
    console.log(`Event mapping updated for ${normalizedUpn}:${outlookEventId}.`);
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
    tryAcquireLock,
    releaseLock,
    findEventMapping,
    findEventMappingByGraphId,
    updateEventMapping,
    deleteEventMapping,
    readSheetRows, // Exposed for runFullSyncForAllUsers might need it
    appendSheetRows,
    batchUpdateSheetRanges,
    clearSheetRange,
};
