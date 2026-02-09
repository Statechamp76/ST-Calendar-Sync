const { DateTime, Interval } = require('luxon');

const TIMEZONE = 'America/Chicago'; // As specified in the requirements

/**
 * Splits a multi-day Outlook event into single-day blocks in the target timezone.
 * Each block represents a non-job appointment for a single day.
 *
 * @param {string} startISO - Start time of the Outlook event in ISO 8601 format (UTC).
 * @param {string} endISO - End time of the Outlook event in ISO 8601 format (UTC).
 * @returns {Array<{start: string, end: string}>} Array of single-day blocks, where 'start' and 'end'
 *          are ISO 8601 strings in the specified TIMEZONE.
 */
function splitMultiDayEvent(startISO, endISO) {
    const start = DateTime.fromISO(startISO, { zone: 'utc' }).setZone(TIMEZONE);
    const end = DateTime.fromISO(endISO, { zone: 'utc' }).setZone(TIMEZONE);

    // If the event starts and ends on the same local day, no splitting is needed.
    if (start.hasSame(end, 'day')) {
        return [{ start: start.toISO(), end: end.toISO() }];
    }

    let blocks = [];
    let cursor = start;

    // Loop through each day the event spans
    while (cursor < end) {
        const currentDayEnd = cursor.endOf('day');

        let blockStart;
        let blockEnd;

        if (cursor.hasSame(start, 'day')) {
            // First day: starts at event start, ends at local day's end
            blockStart = start;
            blockEnd = (end < currentDayEnd) ? end : currentDayEnd;
        } else if (cursor.hasSame(end, 'day')) {
            // Last day: starts at local day's beginning, ends at event end
            blockStart = cursor.startOf('day');
            blockEnd = end;
        } else {
            // Full intermediate day: starts at local day's beginning, ends at local day's end
            blockStart = cursor.startOf('day');
            blockEnd = currentDayEnd;
        }
        
        blocks.push({ start: blockStart.toISO(), end: blockEnd.toISO() });

        // Move cursor to the beginning of the next day
        cursor = currentDayEnd.plus({ milliseconds: 1 }).startOf('day');
    }
    
    return blocks;
}

module.exports = { splitMultiDayEvent, TIMEZONE };
