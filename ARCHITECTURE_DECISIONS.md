# Architecture Decisions

## 2026-02-12: State Migration From Sheets To Firestore

### Decision
Migrate lock, event-map, delta-token, and reconcile bookkeeping state from Google Sheets to Firestore (Native mode), with a feature-flag fallback path.

### Why
- Sheets state experienced quota throttling (`429`) during lock/map/state traffic.
- Sheets row-level locking is not transactional and can produce race conditions.
- Firestore transactions provide stronger compare-and-set semantics for per-mailbox locks.
- Firestore document writes improve idempotency and rollback tracking for create/update flows.

### Scope
- Firestore collections:
  - `locks`
  - `eventMap`
  - `deltaState`
- Sheets retained for:
  - `TechMap` (human editable)
  - optional human-managed keyword/config lists

### Feature Flag
- `FIRESTORE_ENABLED=true`: Firestore state path active.
- `FIRESTORE_ENABLED=false`: legacy Sheets state path remains active for rollback safety.

### Operational Rollback
1. Set `FIRESTORE_ENABLED=false`.
2. Redeploy service.
3. Validate sync cycle and monitor for duplicates via `/reconcile/window` dry-run.

### Additional Policy Choices
- `SYNC_MODE=EXCLUDE_FREE_ONLY`: sync all non-free events.
- Private events are masked with exact title `Private`.
- All-day events are included except birthday/holiday keyword exclusions.
