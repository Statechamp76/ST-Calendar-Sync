# ST Calendar Sync (`st-calendar-sync`)

Collects Outlook (Microsoft 365) calendar events and upserts relevant entries as ServiceTitan non-job appointments.

For continuity between sessions, see `AI_ASSISTANT_NOTES.md` for current operational state and deferred tasks.

## Internal Memory

This repo maintains an "internal memory" file at `AI_ASSISTANT_NOTES.md`.

When doing maintenance or incident response, read that file first for:
- current runtime status
- Sheets schema notes
- idempotency/stable-key invariants
- incident runbooks (maintenance mode, purge, reconcile)

## App Location And Entrypoint

- App root: this repository root (no `outlook-sync-service/` subfolder)
- Entrypoint: `src/index.js`
- Start command: `npm start` (`node src/index.js`)
- Port: `process.env.PORT` with default `8080`

## Endpoints

- `GET /health` -> `200 ok`
- `POST /run-sync` -> triggers one delta sync cycle (all enabled users) and returns JSON summary
- `POST /backfill/last-30-days` -> one-time backfill for past 30 days (busy/OOF only)
- `POST /backfill/next-90-days` -> temporary: backfill a configurable forward window (default 7 days starting 2026-02-14 UTC)
- `POST /cleanup/reset` (OIDC) -> purge ST non-jobs in a window (supports all technicians) and optionally clear sheets
- `POST /cleanup/delete-nonjobs` (OIDC) -> delete explicit ST non-job appointment IDs
- `POST /cleanup/purge-nonjobs-for-tech` (OIDC) -> purge ST non-jobs for one technician ID
- `POST /reconcile/window` (OIDC) -> detect/delete duplicates created by this integration (stable-key based); supports `startUtc`/`endUtc` query or body
- `POST /internal/ai/inbox/process` (OIDC) -> consume `AI_INBOX` rows, write status rows to `AI_OUTBOX`, then clear consumed inbox rows
- `POST /ai/entries` (API key) -> append structured AI memory entry to Firestore
- `GET /ai/entries` (API key) -> read recent AI memory entries from Firestore
- `GET /ai/state` (API key) -> lightweight rollup for a project (counts/latest titles in recent window)

`/run-sync` response shape:

```json
{
  "startedAt": "2026-02-09T17:00:00.000Z",
  "finishedAt": "2026-02-09T17:00:12.125Z",
  "calendarsProcessed": 3,
  "eventsFetched": 42,
  "eventsUpserted": 17,
  "eventsSkipped": 25,
  "errors": []
}
```

## Required Environment Variables

- `RUN_SYNC_AUDIENCE`
- `SYNC_WINDOW_PAST_DAYS` (optional, default `30`)
- `SYNC_WINDOW_FUTURE_DAYS` (optional, default `90`)
- `PUBSUB_TOPIC` (optional, default `outlook-change-notifications`)
- `OUTLOOK_USER_UPNS` (optional comma-separated fallback list)
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`
- `GRAPH_TENANT_ID`
- `GRAPH_WEBHOOK_URL` (required for subscription renewal flow)
- `GRAPH_CLIENT_STATE` (required for subscription renewal flow)
- `SERVICETITAN_CLIENT_ID`
- `SERVICETITAN_CLIENT_SECRET`
- `SERVICETITAN_TENANT_ID`
- `SERVICETITAN_APP_KEY` (optional, if your tenant requires app key header)
- `GOOGLE_SPREADSHEET_ID`
- `ALERT_SLACK_WEBHOOK_URL` (optional)
- `SENDGRID_API_KEY` (optional, for email alerts)
- `ALERT_EMAIL_TO` (optional, requires `SENDGRID_API_KEY`)
- `ALERT_EMAIL_FROM` (optional, requires `SENDGRID_API_KEY`)
- `ALERT_COOLDOWN_SECONDS` (optional, default `600`)
- `ST_CLEAR_DISPATCH_BOARD` (optional, default `true`)
- `ST_CLEAR_TECHNICIAN_VIEW` (optional, default `false`)
- `ST_REMOVE_FROM_CAPACITY` (optional, default `true`)
- `ST_INTEGRATION_APPLICATION_GUID` (optional, default set in code; used for stable-key external data marker)
- `CANONICAL_HUDDLE_MAILBOX` (optional, default `MBrennan@elevatedroofing.com`)
- `SALES_HUDDLE_USER_UPNS` (optional, CSV of Outlook UPNs for huddle targeting)
- `DISABLE_HUDDLE_SYNC` (optional, default `true`; when true, huddle blocks are not created/updated from Outlook)
- `DISABLE_BACKFILL_LAST_30` (optional, default `true`; when true, `POST /backfill/last-30-days` returns 410)
- `BACKFILL_NEXT_START_UTC` (optional, default `2026-02-14T00:00:00Z`)
- `BACKFILL_NEXT_DAYS` (optional, default `7`)
- `AI_BRIDGE_ENABLED` (optional, default `false`; when true, enables `/ai/*` endpoints)
- `AI_BRIDGE_API_KEY` (required when `AI_BRIDGE_ENABLED=true`; must match `X-API-Key` request header)
- `FIRESTORE_PROJECT_ID` (optional; defaults to ADC project)
- `FIRESTORE_ENABLED` (optional, default `false`; when true, state/locks/eventMap/deltaState use Firestore)
- `SYNC_MODE` (optional, default `EXCLUDE_FREE_ONLY`)
- `HOLIDAY_EXCLUDE_KEYWORDS` (optional CSV, default `holiday,vacation`; all-day events with these keywords are excluded)
- `LOCK_TTL_SECONDS` (optional, default `600`)
- `RECONCILE_DAYS_AHEAD` (optional, default `90`)
- `Locks` sheet: this service will auto-create a `Locks` tab in the configured spreadsheet to coordinate per-user sync locks.

## Firestore State Migration

Feature-flagged cutover:
- `FIRESTORE_ENABLED=true`:
  - Locks -> Firestore `locks`
  - Event mapping -> Firestore `eventMap`
  - Delta token state -> Firestore `deltaState`
- `FIRESTORE_ENABLED=false`:
  - fallback to legacy Sheets state path (TechMap remains in Sheets either way)

Why:
- reduce Sheets 429 lock/state failures
- stronger lock semantics via Firestore transactions
- improved idempotency and rollback traceability

Rollback:
1. Set `FIRESTORE_ENABLED=false`
2. Redeploy
3. Run `POST /reconcile/window` in dry-run mode to validate duplicates are controlled

Sync policy:
- Outlook is source of truth (no Outlook writes)
- `SYNC_MODE=EXCLUDE_FREE_ONLY`: sync everything except `showAs=free/available`
- Private events always written to ST with title exactly `Private`
- All-day events are included except:
  - subject contains `birthday`
  - subject matches `HOLIDAY_EXCLUDE_KEYWORDS`

## AI Bridge (Firestore)

Purpose:
- Persist prompts, decisions, runbooks, notes, incidents, and TODO items outside chat.
- Store/retrieve entries under Firestore collection `ai_projects/{project}/entries`.

Security:
- `/ai/*` endpoints are disabled unless `AI_BRIDGE_ENABLED=true`.
- `/ai/*` endpoints require `X-API-Key` matching `AI_BRIDGE_API_KEY`.
- Entry content is never logged; logs only include metadata (`entryId`, `project`, `type`, `title`).
- Guardrails reject payloads containing blocked keys (`attendees`, `body`, `location`) to avoid raw calendar payload dumps.

POST example:

```powershell
curl -X POST "https://<cloud-run-url>/ai/entries" `
  -H "Content-Type: application/json" `
  -H "X-API-Key: <AI_BRIDGE_API_KEY>" `
  --data-binary "@samples/ai_entry.json"
```

GET example:

```powershell
curl "https://<cloud-run-url>/ai/entries?project=ST-Calendar-Sync&limit=20" `
  -H "X-API-Key: <AI_BRIDGE_API_KEY>"
```

## AI Inbox/Outbox (Sheets)

Purpose:
- Consume AI work items from `AI_INBOX` and write processing status/results to `AI_OUTBOX`.
- Immediately clear consumed inbox rows (blank all row cells including `id`) to avoid repeated retries/throttling.

Sheet config used by code:
- Spreadsheet ID: `1DHL_hHvduwxUxm0uAtmCIgolwv6cqDSHUe0hJwW6na8`
- `AI_INBOX` gid `1179786912` (headers: `id, created_utc, project, type, title, tags, content, source`)
- `AI_OUTBOX` gid `1028748998` (headers: `id, created_utc, project, type, title, content, related_inbox_id`)

Behavior of `POST /internal/ai/inbox/process`:
- one read from `AI_INBOX`
- one batched append to `AI_OUTBOX`
- one batched clear operation on consumed inbox rows
- strict order: append outbox first, then clear inbox rows
- malformed inbox rows produce outbox `status=error` entries and are still cleared to prevent retry storms
- uses existing Sheets retry/backoff logic for 429/5xx
- concurrency guard uses Firestore lock doc `system_locks/aiInboxLock` with TTL

Run locally:

1. Start service:
```powershell
npm start
```

2. Get local identity token (audience must match your `RUN_SYNC_AUDIENCE`):
```powershell
$aud = "http://localhost:8080"
$token = gcloud auth print-identity-token --audiences=$aud
```

3. Trigger inbox processing:
```powershell
curl -X POST "http://localhost:8080/internal/ai/inbox/process" `
  -H "Authorization: Bearer $token" `
  -H "Content-Type: application/json" `
  -d "{\"project\":\"ST-Calendar-Sync\",\"limit\":50}"
```

## Local Run

```powershell
npm install
npm start
```

Health check:

```powershell
curl http://localhost:8080/health
```

## Deploy To Cloud Run

Microsoft Graph webhooks require a publicly-reachable URL. This service is deployed with
`--allow-unauthenticated` so Microsoft can call `POST /graph/notifications`, while all internal
worker/admin endpoints remain protected by OIDC (`requireOidcAuth` middleware).

```powershell
gcloud run deploy st-calendar-sync `
  --source . `
  --region us-central1 `
  --service-account st-calendar-sync-sa@<PROJECT_ID>.iam.gserviceaccount.com `
  --allow-unauthenticated
```

Set non-secret env vars:

```powershell
gcloud run services update st-calendar-sync `
  --region us-central1 `
  --update-env-vars RUN_SYNC_AUDIENCE=https://st-calendar-sync-<hash>-uc.a.run.app,SYNC_WINDOW_PAST_DAYS=30,SYNC_WINDOW_FUTURE_DAYS=90
```

Set secrets as env vars:

```powershell
gcloud run services update st-calendar-sync `
  --region us-central1 `
  --set-secrets GRAPH_CLIENT_ID=GRAPH_CLIENT_ID:latest,GRAPH_CLIENT_SECRET=GRAPH_CLIENT_SECRET:latest,GRAPH_TENANT_ID=GRAPH_TENANT_ID:latest,SERVICETITAN_CLIENT_ID=SERVICETITAN_CLIENT_ID:latest,SERVICETITAN_CLIENT_SECRET=SERVICETITAN_CLIENT_SECRET:latest,SERVICETITAN_TENANT_ID=SERVICETITAN_TENANT_ID:latest,GOOGLE_SPREADSHEET_ID=GOOGLE_SPREADSHEET_ID:latest,GRAPH_WEBHOOK_URL=GRAPH_WEBHOOK_URL:latest,GRAPH_CLIENT_STATE=GRAPH_CLIENT_STATE:latest
```

## Cloud Scheduler (OIDC) For `/run-sync`

Give the Scheduler service account permission to invoke Cloud Run:

```powershell
gcloud run services add-iam-policy-binding st-calendar-sync `
  --region us-central1 `
  --member serviceAccount:<SCHEDULER_SA_EMAIL> `
  --role roles/run.invoker
```

Create scheduler job:

```powershell
gcloud scheduler jobs create http st-calendar-sync-job `
  --location us-central1 `
  --schedule "*/15 * * * *" `
  --uri "https://st-calendar-sync-<hash>-uc.a.run.app/run-sync" `
  --http-method POST `
  --oidc-service-account-email "<SCHEDULER_SA_EMAIL>" `
  --oidc-token-audience "https://st-calendar-sync-<hash>-uc.a.run.app"
```

The `--oidc-token-audience` value must match `RUN_SYNC_AUDIENCE`.

## GitHub -> Cloud Run Auto Deploy

Workflow file: `.github/workflows/deploy-cloud-run.yml`

Runs on push to `main` (and manual trigger), then deploys this repo to Cloud Run using Workload Identity Federation.

Configure these GitHub repository settings:

- `Variables`
- `GCP_PROJECT_ID`: your GCP project id
- `CLOUD_RUN_SERVICE`: `st-calendar-sync`
- `CLOUD_RUN_REGION`: `us-central1`

- `Secrets`
- `GCP_WORKLOAD_IDENTITY_PROVIDER`: full provider resource name (`projects/<number>/locations/global/workloadIdentityPools/<pool>/providers/<provider>`)
- `GCP_SERVICE_ACCOUNT`: deployer service account email (example: `github-deployer@<PROJECT_ID>.iam.gserviceaccount.com`)

Required IAM for deployer service account:

- `roles/run.admin`
- `roles/iam.serviceAccountUser` on the runtime service account used by Cloud Run
- `roles/cloudbuild.builds.editor` (for `--source` builds)
- `roles/artifactregistry.writer` (if build artifacts are pushed)
Sync behavior rules:

- Multi-day Outlook events are split into single-day ServiceTitan non-job appointments.
- Delta tombstones (`@removed`) delete previously mapped ServiceTitan non-job appointments.
- Events marked `free`/`available` are not created in ServiceTitan; existing mapped records are removed.
- Events marked `private` are synced to ServiceTitan with the name `Busy`.
- All synced ServiceTitan non-job appointments are created with:
  - "Needs a Timesheet?" unchecked (`timesheetCodeId` is omitted)
  - "Visible to technician schedule in the mobile app" checked (`showOnTechnicianSchedule: true`)
