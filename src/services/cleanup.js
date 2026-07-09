const { DateTime } = require('luxon');
const sheets = require('./sheets');
const servicetitan = require('../api/servicetitan');
const { TIMEZONE } = require('../utils/time');

function parseJsonArray(value) {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function isOurSyncLikeAppointment(appt) {
  // These are the defaults our sync uses. If fields are absent, don't treat as ours.
  if (appt?.showOnTechnicianSchedule !== true) return false;
  if (appt?.clearDispatchBoard !== true) return false;
  if (appt?.clearTechnicianView !== false) return false;
  if (appt?.removeTechnicianFromCapacityPlanning !== true) return false;
  if (appt?.active !== true) return false;

  return true;
}

function makeSignature(appt) {
  return [
    appt?.technicianId ?? '',
    appt?.start ?? '',
    appt?.duration ?? '',
    appt?.name ?? '',
    appt?.allDay ? 'A' : 'T',
    appt?.showOnTechnicianSchedule ? 'S1' : 'S0',
    appt?.clearDispatchBoard ? 'D1' : 'D0',
    appt?.clearTechnicianView ? 'V1' : 'V0',
    appt?.removeTechnicianFromCapacityPlanning ? 'C1' : 'C0',
    appt?.active ? 'X1' : 'X0',
  ].join('|');
}

function makeSlotSignature(appt) {
  // Same technical slot/config but name-agnostic to catch "Busy + detailed" duplicates.
  return [
    appt?.technicianId ?? '',
    appt?.start ?? '',
    appt?.duration ?? '',
    appt?.allDay ? 'A' : 'T',
    appt?.showOnTechnicianSchedule ? 'S1' : 'S0',
    appt?.clearDispatchBoard ? 'D1' : 'D0',
    appt?.clearTechnicianView ? 'V1' : 'V0',
    appt?.removeTechnicianFromCapacityPlanning ? 'C1' : 'C0',
    appt?.active ? 'X1' : 'X0',
  ].join('|');
}

function getNamePriority(appt) {
  const name = String(appt?.name || '').trim();
  if (name === 'Private') return 500; // privacy-safe title must win
  if (name === 'Out of Office') return 300;
  if (name === 'Busy') return 200;
  if (name === 'Sales Huddle') return 150;
  // Detailed/non-generic titles are preferred over "Busy".
  if (name) return 400;
  return 100;
}

function pickEntryToKeep(entries, referencedIds) {
  const referenced = entries.filter((e) => referencedIds.has(e.id));
  const candidates = referenced.length > 0 ? referenced : entries;
  const sorted = [...candidates].sort((a, b) => {
    const p = getNamePriority(b.appt) - getNamePriority(a.appt);
    if (p !== 0) return p;
    return a.id.localeCompare(b.id);
  });
  return sorted[0];
}

async function getReferencedNonJobIdsSet() {
  // EventMap columns: outlook_upn, outlook_event_id, st_nonjob_ids_json, last_hash, last_synced_utc, status
  const rows = await sheets.readSheetRows('EventMap!A2:F');
  const referenced = new Set();
  for (const row of rows) {
    const ids = parseJsonArray(row[2] || '');
    for (const id of ids) {
      if (id) referenced.add(String(id));
    }
  }
  return referenced;
}

function toIsoAtTzDayStart(date, zone) {
  return DateTime.fromISO(date, { zone }).startOf('day').toUTC().toISO();
}

function toIsoAtTzDayEnd(date, zone) {
  return DateTime.fromISO(date, { zone }).endOf('day').toUTC().toISO();
}

function getDefaultStartAndEnd() {
  const now = DateTime.now().setZone(TIMEZONE);
  const start = now.startOf('week'); // Monday in Luxon by default locale; TIMEZONE is consistent.
  const end = now.plus({ days: 90 }).endOf('day');
  return {
    startsOnOrAfter: start.toUTC().toISO(),
    startsOnOrBefore: end.toUTC().toISO(),
  };
}

async function dedupeNonJobsThisWeekForward(options = {}) {
  const {
    startsOnOrAfter = null,
    startsOnOrBefore = null,
    dryRun = true,
  } = options;

  const defaults = getDefaultStartAndEnd();
  const startIso = startsOnOrAfter || defaults.startsOnOrAfter;
  const endIso = startsOnOrBefore || defaults.startsOnOrBefore;

  const referenced = await getReferencedNonJobIdsSet();
  const techMap = await sheets.getTechMap();
  const enabledUsers = techMap.filter((u) => u.enabled && u.st_technician_id);

  const summary = {
    startsOnOrAfter: startIso,
    startsOnOrBefore: endIso,
    techniciansProcessed: 0,
    appointmentsScanned: 0,
    duplicateGroupsFound: 0,
    appointmentsToDelete: 0,
    deleted: 0,
    errors: [],
  };

  for (const user of enabledUsers) {
    summary.techniciansProcessed += 1;
    const techId = String(user.st_technician_id);
    try {
      const appts = await servicetitan.listNonJobs({
        technicianId: techId,
        startsOnOrAfter: startIso,
        startsOnOrBefore: endIso,
        page: 1,
        pageSize: 500,
      });

      const ours = appts.filter(isOurSyncLikeAppointment);
      summary.appointmentsScanned += ours.length;

      const groups = new Map();
      for (const appt of ours) {
        const id = String(appt.id);
        const key = makeSlotSignature(appt);
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ id, appt });
      }

      for (const [, entries] of groups.entries()) {
        if (entries.length <= 1) continue;

        // Keep one best candidate for this slot and delete the rest.
        const keep = pickEntryToKeep(entries, referenced);
        const toDelete = entries.filter((e) => e.id !== keep.id);
        if (toDelete.length === 0) continue;
        summary.duplicateGroupsFound += 1;
        summary.appointmentsToDelete += toDelete.length;
        if (!dryRun) {
          for (const e of toDelete) {
            await servicetitan.deleteNonJob(e.id);
            summary.deleted += 1;
          }
        }
      }
    } catch (error) {
      summary.errors.push({
        technicianId: techId,
        userUpn: user.outlook_upn,
        message: error.message,
      });
    }
  }

  return summary;
}

async function purgeNonJobsInWindow(options = {}) {
  const {
    startsOnOrAfter = null,
    startsOnOrBefore = null,
    dryRun = true,
    includeDisabled = true,
    allTechnicians = false,
  } = options;

  const defaults = getDefaultStartAndEnd();
  const startIso = startsOnOrAfter || defaults.startsOnOrAfter;
  const endIso = startsOnOrBefore || defaults.startsOnOrBefore;

  async function purgeForTechnicianId(technicianIdOrNull) {
    // ServiceTitan frequently caps page sizes; use a conservative size and page until empty/no-new.
    const pageSize = 200;
    let page = 1;
    let totalForTarget = 0;
    const seenIds = new Set();

    while (true) {
      const appts = await servicetitan.listNonJobs({
        technicianId: technicianIdOrNull,
        startsOnOrAfter: startIso,
        startsOnOrBefore: endIso,
        page,
        pageSize,
      });

      if (!appts || appts.length === 0) break;

      let newInPage = 0;
      for (const appt of appts) {
        const id = appt && appt.id !== undefined ? String(appt.id) : null;
        if (!id || seenIds.has(id)) continue;
        seenIds.add(id);
        newInPage += 1;
        totalForTarget += 1;

        summary.appointmentsFound += 1;
        summary.appointmentsToDelete += 1;

        if (!dryRun) {
          await servicetitan.deleteNonJob(id);
          summary.deleted += 1;
        }
      }

      if (newInPage === 0) break; // protect against repeating pages / capped paging quirks
      page += 1;
      if (page > 500) break; // safety cap; global listing could be larger
    }

    return totalForTarget;
  }

  let techIds = [];

  if (allTechnicians) {
    // Best-effort "delete for the whole tenant": list non-jobs without a technicianId filter.
    // Some ServiceTitan tenants don't expose a technicians listing endpoint under dispatch/v2.
    // If the global listing is not supported, we'll fall back to the sheet-driven tech list.
    try {
      const probe = await servicetitan.listNonJobs({
        technicianId: null,
        startsOnOrAfter: startIso,
        startsOnOrBefore: endIso,
        page: 1,
        pageSize: 1,
      });

      if (Array.isArray(probe)) {
        techIds = ['*ALL*'];
      }
    } catch (e) {
      console.warn('cleanup.purge.global_list_not_supported', { message: e.message });
      techIds = [];
    }
  } else {
    const techMap = await sheets.getTechMap();
    techIds = techMap
      .filter((u) => u.st_technician_id && (includeDisabled ? true : Boolean(u.enabled)))
      .map((u) => String(u.st_technician_id));
  }

  if (techIds.length === 0) {
    // Fallback: if allTechnicians is true but global listing isn't supported, use the sheet.
    const techMap = await sheets.getTechMap();
    techIds = techMap
      .filter((u) => u.st_technician_id && (includeDisabled ? true : Boolean(u.enabled)))
      .map((u) => String(u.st_technician_id));
  }

  techIds = [...new Set(techIds)];

  const summary = {
    startsOnOrAfter: startIso,
    startsOnOrBefore: endIso,
    techniciansProcessed: 0,
    techniciansTargeted: techIds.length,
    appointmentsFound: 0,
    appointmentsToDelete: 0,
    deleted: 0,
    errors: [],
  };

  for (const techId of techIds) {
    summary.techniciansProcessed += 1;

    try {
      const isGlobal = techId === '*ALL*';
      const totalForTech = await purgeForTechnicianId(isGlobal ? null : techId);
      console.log('cleanup.purge.tech.complete', { techId: isGlobal ? 'ALL' : techId, totalForTech });
    } catch (error) {
      summary.errors.push({
        technicianId: techId,
        message: error.message,
      });
    }
  }

  return summary;
}

async function resetSyncState(options = {}) {
  const {
    startsOnOrAfter = null,
    startsOnOrBefore = null,
    dryRun = true,
    skipSheetsClear = false,
    includeDisabled = true,
    allTechnicians = false,
  } = options;

  const purgeSummary = await purgeNonJobsInWindow({
    startsOnOrAfter,
    startsOnOrBefore,
    dryRun,
    includeDisabled,
    allTechnicians,
  });

  let sheetsCleared = false;
  const errors = [];
  if (!dryRun && !skipSheetsClear) {
    try {
      // Clear mappings but keep headers.
      await sheets.clearSheetRange('EventMap!A2:F');
      await sheets.clearSheetRange('DeltaState!A2:E');
      sheetsCleared = true;
    } catch (error) {
      errors.push({ message: error.message });
    }
  }

  return {
    dryRun,
    purge: purgeSummary,
    sheetsCleared,
    errors,
  };
}

async function clearSyncSheets() {
  // Clear mappings but keep headers.
  await sheets.clearSheetRange('EventMap!A2:F');
  await sheets.clearSheetRange('DeltaState!A2:E');
  return { cleared: true };
}

async function purgeNonJobsForTechnician(options = {}) {
  const {
    technicianId,
    startsOnOrAfter = null,
    startsOnOrBefore = null,
    dryRun = true,
  } = options;

  const techId = String(technicianId || '').trim();
  if (!techId) throw new Error('Missing required technicianId');

  const defaults = getDefaultStartAndEnd();
  const startIso = startsOnOrAfter || defaults.startsOnOrAfter;
  const endIso = startsOnOrBefore || defaults.startsOnOrBefore;

  const summary = {
    dryRun,
    technicianId: techId,
    startsOnOrAfter: startIso,
    startsOnOrBefore: endIso,
    appointmentsFound: 0,
    deleted: 0,
    errors: [],
  };

  try {
    // ServiceTitan frequently caps page sizes; use a conservative size and page until empty/no-new.
    const pageSize = 200;
    let page = 1;
    const seenIds = new Set();

    while (true) {
      const appts = await servicetitan.listNonJobs({
        technicianId: techId,
        startsOnOrAfter: startIso,
        startsOnOrBefore: endIso,
        page,
        pageSize,
      });

      if (!appts || appts.length === 0) break;

      let newInPage = 0;
      for (const appt of appts) {
        const id = appt && appt.id !== undefined ? String(appt.id) : null;
        if (!id || seenIds.has(id)) continue;
        seenIds.add(id);
        newInPage += 1;
        summary.appointmentsFound += 1;

        if (!dryRun) {
          await servicetitan.deleteNonJob(id);
          summary.deleted += 1;
        }
      }

      if (newInPage === 0) break; // protect against repeating pages / capped paging quirks
      page += 1;
      if (page > 500) break; // safety cap
    }
  } catch (error) {
    summary.errors.push({ message: error.message });
  }

  return summary;
}

async function deleteNonJobsByIds(options = {}) {
  const {
    ids = [],
    dryRun = true,
  } = options;

  const uniqueIds = [...new Set((ids || []).map((x) => String(x || '').trim()).filter(Boolean))];

  const summary = {
    dryRun,
    requested: Array.isArray(ids) ? ids.length : 0,
    unique: uniqueIds.length,
    deleted: 0,
    errors: [],
  };

  for (const id of uniqueIds) {
    try {
      if (!dryRun) {
        await servicetitan.deleteNonJob(id);
        summary.deleted += 1;
      }
    } catch (error) {
      summary.errors.push({ id, message: error.message });
    }
  }

  return summary;
}

module.exports = {
  dedupeNonJobsThisWeekForward,
  purgeNonJobsInWindow,
  purgeNonJobsForTechnician,
  resetSyncState,
  clearSyncSheets,
  deleteNonJobsByIds,
  toIsoAtTzDayStart,
  toIsoAtTzDayEnd,
};
