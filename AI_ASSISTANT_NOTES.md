# AI Assistant Handoff Notes

Last updated: 2026-02-09

## Current Status

- App is deployed on private Cloud Run service `st-calendar-sync` in `us-central1`.
- GitHub Actions auto-deploy is configured (`.github/workflows/deploy-cloud-run.yml`).
- Scheduler job `st-calendar-sync-job` is configured and running.
- `/run-sync` now runs **delta mode** for all enabled users.

## Implemented Sync Rules

- Multi-day Outlook events are split into single-day ServiceTitan non-job appointments.
- Graph tombstones (`@removed`) remove mapped ServiceTitan non-jobs.
- Events marked `free`/`available` are not synced to ServiceTitan (and existing mapped non-jobs are removed).
- Private Outlook events are synced with name `Busy`.

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

## Scheduler / Frequency

- Current schedule is every 15 minutes:
  - `*/15 * * * *`
- If needed, nightly-only can be configured later.

## Known Operational Notes

- Existing appointments already in ServiceTitan do not automatically adopt new payload defaults unless they are updated/recreated.
- Google Sheets quota can be hit under heavy sync volume; current implementation still works but may log quota warnings.
- Deployer SA currently uses broad permissions for stability (least-privilege hardening can be done later).

## Deferred Work (Next Session)

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
