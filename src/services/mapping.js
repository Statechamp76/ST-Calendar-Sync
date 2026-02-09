const { DateTime, Interval } = require('luxon');
const { splitMultiDayEvent, TIMEZONE } = require('../utils/time');

function parseBool(value, defaultValue) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  const normalized = String(value).trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes';
}

function mapEventToServiceTitanPayloads(event, userConfig) {
  const subject = event.isPrivate ? 'Busy' : (event.subject || 'Calendar Event');
  const eventBlocks = splitMultiDayEvent(event.start, event.end);
  const requireTimesheet = parseBool(process.env.ST_REQUIRE_TIMESHEET, false);
  const showOnTechnicianSchedule = parseBool(process.env.ST_SHOW_ON_TECH_SCHEDULE, true);
  const clearDispatchBoard = parseBool(process.env.ST_CLEAR_DISPATCH_BOARD, true);
  const clearTechnicianView = parseBool(process.env.ST_CLEAR_TECHNICIAN_VIEW, false);
  const removeFromCapacity = parseBool(process.env.ST_REMOVE_FROM_CAPACITY, true);

  return eventBlocks.map((block) => {
    const startDateTime = DateTime.fromISO(block.start, { zone: TIMEZONE });
    const endDateTime = DateTime.fromISO(block.end, { zone: TIMEZONE });
    const duration = Interval.fromDateTimes(startDateTime, endDateTime).toDuration().toFormat('hh:mm:ss');

    return {
      technicianId: userConfig.st_technician_id,
      timesheetCodeId: requireTimesheet ? userConfig.st_timesheet_code_id : 0,
      start: startDateTime.toISO(),
      duration,
      name: subject,
      allDay: Boolean(event.isAllDay),
      showOnTechnicianSchedule,
      clearDispatchBoard,
      clearTechnicianView,
      removeTechnicianFromCapacityPlanning: removeFromCapacity,
      active: true,
    };
  });
}

module.exports = {
  mapEventToServiceTitanPayloads,
};
