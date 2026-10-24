# RuPayVault — Pending Tasks & Architecture Roadmap

This document tracks planned architectural enhancements, pending bug fixes, and feature improvements for the RuPayVault project.

---

## 📌 PENDING: Concurrency Guard & Sync-Mutation Interlock ("Shield & Queue")

- **Status**: Pending Planning / Implementation
- **Area**: Frontend Sync (`index.html`) & Backend Consistency (`CronCode.gs`)
- **Reported Issue**: When vouchers are approved to brand tabs (e.g. vouchers tagged with `PNB DC 5151`), clicking "Sync" manually a few seconds later causes them to reappear in the "Ingested" tab temporarily before stabilizing on a second approval.

### 1. Objective
Eliminate read-after-write replication race conditions between client-side optimistic UI updates, background mutating `apiFetch` calls (such as `APPROVE_STAGING`, `MOVE_TO_REDEEMED`, `TAG_CARD`), and manual or automated `syncWithGoogleSheets()` (`doGet`) fetches from Google Sheets.

### 2. Root Cause Analysis
1. **Optimistic UI vs. Google Apps Script Latency**:
   - The UI immediately updates local memory (`state.vouchers`, `state.staging`) and presents an instant success state.
   - Background `POST` requests to Google Apps Script take 3–8 seconds to acquire script locks, append rows to destination tabs, invoke `newSheet.deleteRow()`, and commit via `SpreadsheetApp.flush()`.
2. **Premature Read via Manual Sync**:
   - `syncWithGoogleSheets()` sends a `GET` request without verifying if background mutations are still in flight.
   - `doGet` reads the sheet state during the window where rows are appended to the brand tab but not yet deleted/propagated from the `"New"` staging tab.
3. **Blind Local Overwrite**:
   - `syncWithGoogleSheets()` blindly overwrites `state.staging` with `data.staging` returned by the server, resurrecting staged items locally.

---

### 3. Suggested Solutions (3-Pillar "Shield & Queue" Strategy)

#### Pillar 1: Global Mutation Mutex & Visual Locking
- Introduce a reactive mutation counter: `state.pendingMutationsCount`.
- Increment when any mutating `apiFetch` begins (`APPROVE_STAGING`, `MOVE_TO_REDEEMED`, `UNDO_REDEEM`, `TAG_CARD`, `CLEAR_ACTIVITY`, etc.) and decrement upon completion (in a `finally` block).
- While `state.pendingMutationsCount > 0`:
  - Disable manual sync triggers (header sync button, drawer sync item).
  - Replace the sync icon with an active "Saving to Google Sheets..." pulse indicator.
  - If the user clicks Sync while locked, show a subtle toast: *"Sync queued — finishing current changes first"*.
- **Fail-Safe Timeout**: Include an automatic 15-second expiration on any lock to ensure the UI never permanently freezes if a network request drops.

#### Pillar 2: Client-Side Quarantine / Tombstone Filter (Shield)
- Track in-flight or recently mutated identifiers in an ephemeral Set with a time-to-live (TTL, e.g., 20–30 seconds):
  ```javascript
  state.quarantinedStagingIds = new Map(); // bookingId/code -> expirationTimestamp
  ```
- When `APPROVE_STAGING` or `CLEAR_ACTIVITY` executes, record the target IDs into `quarantinedStagingIds`.
- When `syncWithGoogleSheets()` receives `data.staging` from the server, filter out any entries present in the active quarantine:
  ```javascript
  const now = Date.now();
  state.staging = data.staging.filter(s => {
    const expiry = state.quarantinedStagingIds.get(s.bookingId) || state.quarantinedStagingIds.get(s.code);
    return !expiry || expiry < now;
  });
  ```
- **Benefit**: Even if Google Sheets' replication lags by 2–3 seconds after an HTTP `200 OK`, the server snapshot *cannot* resurrect items in the UI.

#### Pillar 3: Sync Queueing (Deferred Auto-Sync)
- Instead of silently dropping user-initiated or automated sync requests during an active operation, set a flag: `state.syncPendingAfterMutation = true`.
- Once `state.pendingMutationsCount` returns to 0 and all in-flight mutations resolve, automatically execute the queued sync cleanly without user intervention.

---

### 4. Operations Covered by this Architecture
- `APPROVE_STAGING` (Staging $\rightarrow$ Brand tabs)
- `MOVE_TO_REDEEMED` / `MARK_REDEEMED` (Brand tabs $\rightarrow$ Redeemed)
- `UNDO_REDEEM` (Redeemed $\rightarrow$ Brand tabs)
- `TAG_CARD` (Card variant & last-4 assignment)
- `CLEAR_ACTIVITY` (Dismiss staging or shared items)
- `CLEAR_REDEEMED` (Dismiss redeemed history items)
- `SAVE_CARD`, `BLOCK_CARD`, `UPDATE_CARD_QUOTA` (Cards tab mutations)
