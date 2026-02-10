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
  const name = String(appt?.name || '').trim();
  if (!(name === 'Busy' || name === 'Out of Office')) return false;

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
        const key = makeSignature(appt);
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ id, appt });
      }

      for (const [, entries] of groups.entries()) {
        if (entries.length <= 1) continue;

        // If at least one of these IDs is referenced in EventMap, delete only the unreferenced duplicates.
        // Otherwise keep one and delete the rest.
        const referencedEntries = entries.filter((e) => referenced.has(e.id));
        const unref = entries.filter((e) => !referenced.has(e.id));

        if (referencedEntries.length > 0) {
          if (unref.length === 0) continue;
          summary.duplicateGroupsFound += 1;
          summary.appointmentsToDelete += unref.length;
          if (!dryRun) {
            for (const e of unref) {
              await servicetitan.deleteNonJob(e.id);
              summary.deleted += 1;
            }
          }
          continue;
        }

        // No referenced IDs: keep one (lowest id) and delete the rest.
        const sorted = [...entries].sort((a, b) => a.id.localeCompare(b.id));
        const toDelete = sorted.slice(1);
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

  let techIds = [];

  if (allTechnicians) {
    // Prefer ServiceTitan as the source of truth for the purge target list.
    // This is the only reliable way to delete *all* non-job appointments across the tenant.
    const all = [];
    let page = 1;
    const pageSize = 500;
    while (true) {
      const batch = await servicetitan.listTechnicians({ page, pageSize });
      all.push(...batch);
      if (!batch || batch.length < pageSize) break;
      page += 1;
      if (page > 200) break; // safety cap
    }
    techIds = all
      .map((t) => t && (t.id ?? t.technicianId))
      .filter((id) => id !== undefined && id !== null && String(id).trim() !== '')
      .map((id) => String(id));
  } else {
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
      const pageSize = 500;
      let page = 1;
      let totalForTech = 0;
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

        for (const appt of appts) {
          const id = appt && appt.id !== undefined ? String(appt.id) : null;
          if (!id || seenIds.has(id)) continue;
          seenIds.add(id);
          totalForTech += 1;

          summary.appointmentsFound += 1;
          summary.appointmentsToDelete += 1;

          if (!dryRun) {
            await servicetitan.deleteNonJob(id);
            summary.deleted += 1;
          }
        }

        if (appts.length < pageSize) break;
        page += 1;
        if (page > 200) break; // safety cap
      }

      console.log('cleanup.purge.tech.complete', { techId, totalForTech });
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

module.exports = {
  dedupeNonJobsThisWeekForward,
  purgeNonJobsInWindow,
  resetSyncState,
  clearSyncSheets,
  toIsoAtTzDayStart,
  toIsoAtTzDayEnd,
};
