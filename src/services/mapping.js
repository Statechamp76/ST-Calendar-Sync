const { DateTime, Interval } = require('luxon');
const { splitMultiDayEvent, TIMEZONE } = require('../utils/time');

function mapEventToServiceTitanPayloads(event, userConfig) {
  const subject = event.isPrivate ? 'Busy' : (event.subject || 'Calendar Event');
  const eventBlocks = splitMultiDayEvent(event.start, event.end);

  return eventBlocks.map((block) => {
    const startDateTime = DateTime.fromISO(block.start, { zone: TIMEZONE });
    const endDateTime = DateTime.fromISO(block.end, { zone: TIMEZONE });
    const duration = Interval.fromDateTimes(startDateTime, endDateTime).toDuration().toFormat('hh:mm:ss');

    return {
      technicianId: userConfig.st_technician_id,
      timesheetCodeId: userConfig.st_timesheet_code_id,
      start: startDateTime.toISO(),
      duration,
      name: subject,
    };
  });
}

module.exports = {
  mapEventToServiceTitanPayloads,
};
