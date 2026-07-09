# AI Assistant Notes (Internal Memory)

Last updated: 2026-02-12 (Firestore migration prep)

This file is the project's "internal memory". When asked about current status, incident runbooks, or how syncing works, consult this file first.

## 2026-02-12 Update: AI Inbox/Outbox Processor

What changed:
- Added `src/services/aiInbox.js` to process Google Sheets AI inbox/outbox flow.
- Added internal OIDC-protected endpoint:
  - `POST /internal/ai/inbox/process`
- Processing flow is intentionally minimal:
  1. one inbox read (`AI_INBOX!A2:H`)
  2. one batched outbox append (`AI_OUTBOX!A:G`)
  3. one batched inbox clear (blank `A:H` for consumed rows, including `id`)
- Added batched helpers in `src/services/sheets.js`:
  - `appendSheetRows`
  - `batchUpdateSheetRanges`
- Reused existing Sheets retry/backoff for 429/5xx (`withRetry`).
- Added tests:
  - `test/aiInbox.test.js` (read parse + clear ranges)

Operational notes:
- Spreadsheet ID is pinned in code for inbox/outbox module:
  - `1DHL_hHvduwxUxm0uAtmCIgolwv6cqDSHUe0hJwW6na8`
- If a row cannot be processed cleanly, the processor still emits an outbox error/status row and clears the inbox row to prevent retry storms.
- Concurrency guard is enabled via Firestore TTL lock document:
  - `system_locks/aiInboxLock`
  - competing runs return `skipped=true, reason=locked` (endpoint returns HTTP 409).

Rollback:
1. Disable scheduler/job invoking `/internal/ai/inbox/process` (if configured).
2. Revert `src/services/aiInbox.js` mount in `src/app.js` and batched helpers in `src/services/sheets.js`.
3. No data migration is required; cleared inbox rows remain blank by design.

## 2026-02-12 Update: Sheets -> Firestore State Migration

Decision:
- Runtime state moved behind a new state abstraction with Firestore primary path and Sheets fallback.
- Human-managed config remains in Sheets (`TechMap`).

Why:
- Existing incidents showed Sheets 429/read-lock fragility.
- Firestore transactions provide stronger lock semantics and reduce duplicate risk from races.

What changed:
- Added `src/services/stateStore.js` with feature-flagged state ops:
  - lock acquire/release
  - event map get/upsert/delete/rolled-back markers
  - delta token get/set
- Added/kept Firestore collections:
  - `locks`
  - `eventMap`
  - `deltaState`
- Sync now uses `stateStore` for locks + mapping + delta state.
- Pub/Sub publishes now include ordering key per mailbox (`orderingKey=userUpn`).
- Sync policy aligned:
  - `SYNC_MODE=EXCLUDE_FREE_ONLY`
  - private title forced to `Private`
  - all-day excludes for birthday/holiday keywords
- Graph API now has retry/backoff on 429/5xx + transient network failures.

Required env vars:
- `FIRESTORE_ENABLED=true` (or `false` for legacy fallback)
- `FIRESTORE_PROJECT_ID` (optional)
- `SYNC_MODE=EXCLUDE_FREE_ONLY`
- `HOLIDAY_EXCLUDE_KEYWORDS=holiday,vacation`
- `LOCK_TTL_SECONDS=600`
- `RECONCILE_DAYS_AHEAD=90`

Rollback plan:
1. Set `FIRESTORE_ENABLED=false`.
2. Redeploy.
3. Validate with `/run-sync` and `/reconcile/window` dry-run.
4. Keep `MAINTENANCE_MODE=true` during aggressive purge/rebuild incidents.

Known incident context:
- Duplicates were historically tied to mapping/lock persistence races and Sheets 429 lockouts.
- Firestore lock + mapping path is intended to reduce this class of incidents.

## 2026-02-12 Update: AI Bridge (Firestore)

What changed:
- Added API-key protected AI Bridge routes under `/ai`:
  - `POST /ai/entries`
  - `GET /ai/entries`
  - `GET /ai/state`
- Added Firestore-backed storage:
  - `src/services/firestore.js` (firebase-admin + ADC)
  - `src/services/aiBridgeStore.js` (project upsert + entry add/read)
- Added validation/safety guardrails:
  - required fields + type allow-list + content length cap (50,000)
  - blocked key detection for `attendees`, `body`, `location` (deep scan)
  - light content-pattern rejection for obvious raw payload field dumps
- Added tests for API key and payload validation (`test/aiBridge.test.js`).

How to run:
1. Set env vars in Cloud Run:
   - `AI_BRIDGE_ENABLED=true`
   - `AI_BRIDGE_API_KEY=<secret>`
   - optional `FIRESTORE_PROJECT_ID=<gcp-project-id>`
2. Ensure runtime service account has Firestore access (Datastore User or equivalent for Native Firestore reads/writes).
3. Call `/ai/entries` with `X-API-Key` and structured JSON payload.

Rollback plan:
1. Immediate disable: set `AI_BRIDGE_ENABLED=false` and redeploy/update service env vars.
2. If needed, remove router mount (`/ai`) from `src/app.js` in the next code deploy.
3. Firestore data is append-only by design; rollback does not require data migration.

## Current Status

- App is deployed on private Cloud Run service `st-calendar-sync` in `us-central1`.
- GitHub Actions auto-deploy is configured (`.github/workflows/deploy-cloud-run.yml`).
- Scheduler jobs:
  - `st-calendar-sync-job` (*/15) is ENABLED (normal cadence)
  - `st-calendar-sync-graph-renew` (every 6 hours) is ENABLED (keeps Graph subscriptions renewed)
- `/run-sync` now runs **delta mode** for all enabled users.
- Incident tooling endpoints exist for aggressive cleanup and targeted deletes (ServiceTitan only).
- `MAINTENANCE_MODE` is now set to `false` in Cloud Run (sync is live again).
- `DISABLE_BACKFILL_LAST_30=true` (endpoint disabled).
- Current backfill window config:
  - `BACKFILL_NEXT_START_UTC=2026-02-14T00:00:00Z`
  - `BACKFILL_NEXT_DAYS=90`
- Per-user distributed locking is enabled via a Sheets tab `Locks` (auto-created).

## TODO

- [x] Confirm ServiceTitan integration stable key persistence (verified by read-back on create/update)
- [x] Add per-user concurrency lock (Sheets-based TTL lock) to reduce races across workers
- [ ] Implement true "single ST record for all sales techs" if ServiceTitan supports multi-tech non-jobs
- [ ] Add unit tests for MAINTENANCE_MODE endpoint gating and rollback behavior under mapping write failure
- [ ] Add a daily Scheduler job for `/reconcile/window` (document only; keep paused by default)

## Implemented Sync Rules

- Multi-day Outlook events are split into single-day ServiceTitan non-job appointments.
- Graph tombstones (`@removed`) remove mapped ServiceTitan non-jobs.
- Events marked `free`/`available` are not synced to ServiceTitan (and existing mapped non-jobs are removed).
- Private Outlook events are synced with name `Busy`.
- Events marked `tentative` are synced (treated as blocking).
- "Sales Huddle" patterns are treated specially (see **Huddles** below).

## ServiceTitan Payload Behavior

- Default mapping now targets mobile-visible blocking behavior:
  - `timesheetCodeId` omitted (always; "Needs a Timesheet?" unchecked)
  - `showOnTechnicianSchedule: true` (always; visible in mobile tech schedule)
  - `clearDispatchBoard: true`
  - `clearTechnicianView: false`
  - `removeTechnicianFromCapacityPlanning: true`
  - `active: true`

Environment toggles:

- `ST_CLEAR_DISPATCH_BOARD` (default `true`)
- `ST_CLEAR_TECHNICIAN_VIEW` (default `false`)
- `ST_REMOVE_FROM_CAPACITY` (default `true`)

## Idempotency (Stable Key + External Data)

Goal: prevent duplicates even if Sheets mapping is missing or Graph IDs change.

- For every ST non-job create/update, we attach an integration marker with:
  - `applicationGuid` (env `ST_INTEGRATION_APPLICATION_GUID`)
  - `externalId` (stable key string)
- Stable key format (normal events):
  - `{tenantId}:{technicianId}:{iCalUId|graphId}:{startUtcIso}:{endUtcIso}`
- Logs never print the full stable key; they print `stableKeyHash` (sha256 prefix).

Notes:
- ServiceTitan external data payload shape varies; code tries several common shapes on create/update.
- ExternalData is verified by read-back:
  - create/update reads `GET /non-job-appointments/{id}` and confirms the stable key persisted
  - if create verification fails, the created ST record is deleted immediately (rollback)

## Huddles (Canonical Mailbox Router)

Requirement: avoid per-mailbox duplicates for recurring huddles.

- Default behavior: huddles are managed directly in ServiceTitan, so the sync should NOT create/update them.
  - env `DISABLE_HUDDLE_SYNC` (default `true`)
  - when true, the service only cleans up any previously integration-created huddle artifacts

- Huddle events are detected by local-time slot or subject "Sales Huddle":
  - Monday 08:30-09:00 (America/Chicago)
  - Thursday 08:30-09:00 (America/Chicago)
  - Tuesday 10:00-12:00 (America/Chicago)
- Only the canonical mailbox is allowed to drive huddle creation:
  - env `CANONICAL_HUDDLE_MAILBOX` (default `MBrennan@elevatedroofing.com`)
- For non-canonical mailboxes:
  - huddle events are skipped
  - best-effort cleanup runs (delete mapped huddle blocks and also delete by stable key in ST)
- For canonical mailbox:
  - huddles are upserted for the configured sales group:
    - env `SALES_HUDDLE_USER_UPNS` (CSV of Outlook UPNs)
  - if not configured, defaults to just the canonical mailbox.

Current limitation:
- Multi-technician assignment to a single ST non-job appointment is not implemented yet. Canonical huddles are currently created per technician in the configured list, but only from the canonical mailbox (so no dupes from other mailboxes).

## Sheets Schema Notes

TechMap `TechMap!A:D`:
- `outlook_upn`
- `st_technician_id`
- `st_timesheet_code_id` (unused for non-jobs; timesheet is always disabled)
- `enabled` (TRUE/FALSE)

EventMap `EventMap!A:F`:
- `outlook_upn`
- `outlook_event_id` (stable occurrence key like `iCalUId:start:end`)
- `st_nonjob_ids_json` (JSON array of ST IDs)
- `last_hash` (dedupe key)
- `last_synced_utc`
- `status` (includes marker `gid=<graphId>` for tombstone lookup)

DeltaState `DeltaState!A:E`:
- `outlook_upn`
- `delta_link`
- `window_end` (placeholder)
- `last_run_utc`

## Incident Runbooks

### Maintenance Mode

Set `MAINTENANCE_MODE=true` to stop creation while allowing cleanup:
- `/graph/notifications` accepts but does not publish
- `/sync/user` no-ops (204)

### Purge (ServiceTitan Only)

Purge means: delete ServiceTitan Non-Job Appointments only. Never touch Outlook.

Endpoints (OIDC protected):
- `POST /cleanup/reset` (wide purge window; can target all techs)
- `POST /cleanup/purge-nonjobs-for-tech` (purge one techId)
- `POST /cleanup/delete-nonjobs` (delete explicit IDs)

### Reconcile Duplicates (Integration Only)

Endpoint (OIDC protected):
- `POST /reconcile/window`

Behavior:
- lists ST non-jobs in a window (global list if supported; otherwise per-tech via TechMap)
- loads details and extracts integration stable key
- deletes duplicates (keeps newest-ish id)

## Scheduler / Frequency

- Current schedule is every 15 minutes:
  - `*/15 * * * *`
- If needed, nightly-only can be configured later.

## Known Operational Notes

- Existing appointments already in ServiceTitan do not automatically adopt new payload defaults unless they are updated/recreated.
- Google Sheets quota can be hit under heavy sync volume; current implementation still works but may log quota warnings.
- Deployer SA currently uses broad permissions for stability (least-privilege hardening can be done later).

## Deferred Work (Next Session)

### 0) Multi-tech Non-Job Appointment Capability Check

- Determine whether ServiceTitan non-job appointments support assigning multiple technician IDs.
- If supported, implement true huddle collapsing to a single ST record for the sales group.

### 1) Notifications (Slack + Email)

Code is already in place (`src/services/alerts.js`) but runtime config is not finalized.

Supported env/secrets:

- `ALERT_SLACK_WEBHOOK_URL`
- `SENDGRID_API_KEY`
- `ALERT_EMAIL_TO`
- `ALERT_EMAIL_FROM`
- `ALERT_COOLDOWN_SECONDS` (default `600`)

Suggested next steps:

1. Add Slack webhook secret and SendGrid API key secret in GCP.
2. Bind both secrets to Cloud Run service envs.
3. Set `ALERT_EMAIL_TO` and `ALERT_EMAIL_FROM`.
4. Force a test failure (or add a temporary test endpoint) and verify alerts.

### 2) Optional Cleanup / Hardening

1. Tighten GitHub deployer IAM from broad role to least privilege.
2. Reduce Sheets read pressure with caching/batching in `sheets.js`.
3. Add targeted rebuild script for existing ST records if payload normalization changes again.

## Useful Commands

Run sync now:

```powershell
gcloud scheduler jobs run st-calendar-sync-job --location us-central1
```

Read recent logs:

```powershell
gcloud run services logs read st-calendar-sync --region us-central1 --freshness=10m --limit 200
```

Check scheduler config:

```powershell
gcloud scheduler jobs describe st-calendar-sync-job --location us-central1
```
