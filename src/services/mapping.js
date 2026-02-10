const { DateTime, Interval } = require('luxon');
const { splitMultiDayEvent, TIMEZONE } = require('../utils/time');

function mapEventToServiceTitanPayloads(event, userConfig) {
  const subject = event.isPrivate ? 'Busy' : (event.subject || 'Calendar Event');
  const eventBlocks = splitMultiDayEvent(event.start, event.end);

  // Policy: Outlook is the source of truth; these blocks are always non-timesheet and always visible
  // on the technician mobile schedule.
  const showOnTechnicianSchedule = true;
  // NOTE: Some ServiceTitan tenants reject a numeric 0 timesheet code. To keep "Needs a Timesheet?"
  // unchecked, we omit `timesheetCodeId` entirely.

  // Keep these as configuration knobs; they affect where the blocks show up in ST.
  const clearDispatchBoard = (String(process.env.ST_CLEAR_DISPATCH_BOARD || '').trim().toLowerCase() === 'false') ? false : true;
  const clearTechnicianView = (String(process.env.ST_CLEAR_TECHNICIAN_VIEW || '').trim().toLowerCase() === 'true') ? true : false;
  const removeFromCapacity = (String(process.env.ST_REMOVE_FROM_CAPACITY || '').trim().toLowerCase() === 'false') ? false : true;

  return eventBlocks.map((block) => {
    const startDateTime = DateTime.fromISO(block.start, { zone: TIMEZONE });
    const endDateTime = DateTime.fromISO(block.end, { zone: TIMEZONE });
    const duration = Interval.fromDateTimes(startDateTime, endDateTime).toDuration().toFormat('hh:mm:ss');

    return {
      technicianId: userConfig.st_technician_id,
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
