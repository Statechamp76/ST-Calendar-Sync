const servicetitan = require('../api/servicetitan');
const sheets = require('./sheets');
const stateStore = require('./stateStore');
const { loadConfig } = require('../config');
const { extractIntegrationExternalId } = require('../utils/externalData');

const config = loadConfig();

function normalizeId(x) {
  const s = String(x || '').trim();
  return s ? s : null;
}

async function listNonJobsPaged(options = {}) {
  const {
    technicianId = null,
    startsOnOrAfter,
    startsOnOrBefore,
  } = options;

  const all = [];
  const pageSize = 200;
  let page = 1;
  const seenIds = new Set();

  while (true) {
    const batch = await servicetitan.listNonJobs({
      technicianId,
      startsOnOrAfter,
      startsOnOrBefore,
      page,
      pageSize,
    });
    if (!batch || batch.length === 0) break;

    let newInPage = 0;
    for (const appt of batch) {
      const id = normalizeId(appt && appt.id);
      if (!id || seenIds.has(id)) continue;
      seenIds.add(id);
      newInPage += 1;
      all.push(appt);
    }

    if (newInPage === 0) break;
    page += 1;
    if (page > 500) break;
  }

  return all;
}

async function reconcileWindow(options = {}) {
  const {
    startsOnOrAfter,
    startsOnOrBefore,
    dryRun = true,
  } = options;

  if (!startsOnOrAfter || !startsOnOrBefore) {
    throw new Error('Missing required startsOnOrAfter/startsOnOrBefore');
  }

  const summary = {
    dryRun,
    startsOnOrAfter,
    startsOnOrBefore,
    mode: 'global_list_then_detail',
    appointmentsScanned: 0,
    integrationAppointments: 0,
    duplicateGroupsFound: 0,
    appointmentsToDelete: 0,
    deleted: 0,
    errors: [],
  };

  let appts = null;
  try {
    appts = await listNonJobsPaged({
      technicianId: null,
      startsOnOrAfter,
      startsOnOrBefore,
    });
  } catch (e) {
    // Some tenants do not support global listing; fall back to TechMap technician IDs.
    summary.mode = 'per_tech_list_then_detail';
    const techMap = await sheets.getTechMap();
    const techIds = [...new Set(techMap.filter((u) => u.st_technician_id).map((u) => String(u.st_technician_id)))];
    appts = [];
    for (const techId of techIds) {
      try {
        const forTech = await listNonJobsPaged({
          technicianId: techId,
          startsOnOrAfter,
          startsOnOrBefore,
        });
        appts.push(...forTech);
      } catch (err) {
        summary.errors.push({ technicianId: techId, message: err.message });
      }
    }
  }

  const byStableKey = new Map(); // stableKey -> [{id, detail}]

  for (const appt of appts || []) {
    const id = normalizeId(appt && appt.id);
    if (!id) continue;
    summary.appointmentsScanned += 1;

    try {
      const detail = await servicetitan.getNonJob(id);
      const stableKey = extractIntegrationExternalId(detail, config.stIntegrationApplicationGuid);
      if (!stableKey) continue;

      summary.integrationAppointments += 1;
      if (!byStableKey.has(stableKey)) byStableKey.set(stableKey, []);
      byStableKey.get(stableKey).push({ id, detail });
    } catch (e) {
      summary.errors.push({ appointmentId: id, message: e.message });
    }
  }

  for (const [stableKey, entries] of byStableKey.entries()) {
    if (entries.length <= 1) continue;
    summary.duplicateGroupsFound += 1;

    // Keep the newest-ish entry (highest numeric id if possible), delete the rest.
    const sorted = [...entries].sort((a, b) => {
      const an = Number.parseInt(a.id, 10);
      const bn = Number.parseInt(b.id, 10);
      if (Number.isFinite(an) && Number.isFinite(bn)) return bn - an;
      return String(b.id).localeCompare(String(a.id));
    });
    const toDelete = sorted.slice(1);
    summary.appointmentsToDelete += toDelete.length;

    if (!dryRun) {
      for (const e of toDelete) {
        try {
          await servicetitan.deleteNonJob(e.id);
          summary.deleted += 1;
        } catch (err) {
          summary.errors.push({ appointmentId: e.id, message: err.message });
        }
      }
    }
  }

  // Reconcile bookkeeping marker (Firestore path only).
  if (stateStore.isFirestoreEnabled()) {
    try {
      const techMap = await sheets.getTechMap();
      for (const u of techMap.filter((x) => x.enabled && x.outlook_upn)) {
        // Re-anchor the delta window. A Graph calendarView/delta token bakes its
        // [now-pastDays, now+futureDays] window in at initialization and NEVER
        // slides; the token rotates each cycle but the window stays fixed. Once
        // wall-clock time passes the original endDateTime, all current/future
        // events fall outside it and are never returned, so calendars silently
        // stop syncing (root cause of the 2026-07 outage). Clearing the token
        // here forces the next delta cycle to re-initialize a fresh window, so
        // it can never be more than ~24h stale. Idempotent: the next full pull
        // upserts into the same EventMap keys (iCalUId:start:end) rather than
        // duplicating, and this reconcile pass dedups any residual copies.
        await stateStore.setDeltaState(u.outlook_upn, null);
        await stateStore.setLastFullReconcileUtc(u.outlook_upn);
      }
    } catch (e) {
      summary.errors.push({ message: `reconcile_bookkeeping_failed: ${e.message}` });
    }
  }

  return summary;
}

module.exports = {
  reconcileWindow,
};
