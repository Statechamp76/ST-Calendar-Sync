# ST Calendar Sync (`st-calendar-sync`)

Collects Outlook (Microsoft 365) calendar events and upserts relevant entries as ServiceTitan non-job appointments.

## App Location And Entrypoint

- App root: this repository root (no `outlook-sync-service/` subfolder)
- Entrypoint: `src/index.js`
- Start command: `npm start` (`node src/index.js`)
- Port: `process.env.PORT` with default `8080`

## Endpoints

- `GET /health` -> `200 ok`
- `POST /run-sync` -> triggers one delta sync cycle (all enabled users) and returns JSON summary

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

## Local Run

```powershell
npm install
npm start
```

Health check:

```powershell
curl http://localhost:8080/health
```

## Deploy To Cloud Run (Private)

```powershell
gcloud run deploy st-calendar-sync `
  --source . `
  --region us-central1 `
  --service-account st-calendar-sync-sa@<PROJECT_ID>.iam.gserviceaccount.com `
  --no-allow-unauthenticated
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
