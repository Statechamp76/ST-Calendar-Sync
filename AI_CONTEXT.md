# AI Context

## Mission
Synchronize Outlook calendar events into ServiceTitan non-job appointments with high idempotency, minimal duplicates, and safe incident controls.

## Source Of Truth
- Outlook (Microsoft Graph) is the only source of truth for event state.
- The service must never mutate Outlook events.

## Current State Direction
- Human-managed config remains in Google Sheets (`TechMap`, optional keyword/config tabs).
- Fragile runtime state is migrating to Firestore:
  - per-mailbox locks
  - per-event mapping state
  - per-mailbox delta tokens
  - reconcile bookkeeping markers

## Key Runtime Controls
- `MAINTENANCE_MODE=true` pauses sync mutation paths but keeps cleanup endpoints available.
- `FIRESTORE_ENABLED=true` enables Firestore state path; `false` falls back to legacy Sheets state.

## Incident Notes
- Known historical incident pattern: duplicate ST non-jobs due to race/mapping persistence gaps.
- Known scaling issue: Sheets read quota (`429`) under frequent lock/map/state reads.
