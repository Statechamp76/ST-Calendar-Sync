# ST Calendar Sync (`st-calendar-sync`)

Collects Outlook (Microsoft 365) calendar events and upserts relevant entries as ServiceTitan non-job appointments.

## App Location And Entrypoint

- App root: this repository root (no `outlook-sync-service/` subfolder)
- Entrypoint: `src/index.js`
- Start command: `npm start` (`node src/index.js`)
- Port: `process.env.PORT` with default `8080`

## Endpoints

- `GET /health` -> `200 ok`
- `POST /run-sync` -> triggers one sync cycle and returns JSON summary

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
