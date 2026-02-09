const test = require('node:test');
const assert = require('node:assert/strict');
const { mapEventToServiceTitanPayloads } = require('../src/services/mapping');

test('mapEventToServiceTitanPayloads maps single block event', () => {
  const userConfig = {
    st_technician_id: '100',
    st_timesheet_code_id: '200',
  };
  const event = {
    subject: 'Maintenance',
    isPrivate: false,
    start: '2026-02-10T16:00:00.000Z',
    end: '2026-02-10T17:30:00.000Z',
  };

  const payloads = mapEventToServiceTitanPayloads(event, userConfig);

  assert.equal(payloads.length, 1);
  assert.equal(payloads[0].technicianId, '100');
  assert.equal(payloads[0].timesheetCodeId, '200');
  assert.equal(payloads[0].duration, '01:30:00');
  assert.equal(payloads[0].name, 'Maintenance');
});

test('mapEventToServiceTitanPayloads masks private event subject', () => {
  const userConfig = {
    st_technician_id: '100',
    st_timesheet_code_id: '200',
  };
  const event = {
    subject: 'Confidential',
    isPrivate: true,
    start: '2026-02-10T16:00:00.000Z',
    end: '2026-02-10T16:30:00.000Z',
  };

  const payloads = mapEventToServiceTitanPayloads(event, userConfig);
  assert.equal(payloads[0].name, 'Busy');
});
