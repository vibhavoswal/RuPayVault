# RuPayVault UI & Design System Rules

## 1. Zero Unicode Emojis
- **NEVER use Unicode emojis** (e.g. 📥, ✓, 💳, 🏦, 🗂️, ⚡, 🔍, ⚙️, 🔒, etc.) anywhere in the application interface, modals, dropdowns, headers, badges, or buttons.
- All visual glyphs, status markers, and category identifiers must use standard **Google Material Design SVG icons** (or equivalent clean vector SVGs).

## 2. Google Material Design SVG Icons Only
- Use clean, accessible, inline or symbol Google Material Design / Material Symbols SVG icons.
- SVGs must have:
  - Explicit `viewBox="0 0 24 24"`
  - Scalable sizing (`w-4 h-4`, `w-5 h-5`, etc.)
  - Explicit fill or stroke styling (`fill="currentColor"` or `stroke="currentColor"`)
  - Semantic ARIA attributes (e.g., `aria-hidden="true"`)

## 3. Light & Dark Mode Parity
- Every UI element, card, popover, modal, floating dock, and dropdown must look intentionally designed with proper contrast in both Light Mode (`data-theme="light"` / `html.light`) and Dark Mode (`data-theme="dark"` / `html.dark`).
- Avoid hardcoded dark values (e.g. `bg-neutral-800`, `bg-neutral-900`, `text-white`) inside popovers without light mode adaptors or CSS variables.
- In light mode:
  - Surface backgrounds should be clean whites, neutral slates (`#f8fafc`, `#ffffff`), or subtle tinted panels.
  - Active/selected states must use soft tinted backgrounds (e.g., emerald-50 with emerald-200 border and emerald-800 text), never harsh opaque dark blocks.
  - Borders should be subtle (`border-slate-200`), text should be high-contrast (`text-slate-800`, `text-slate-900`).

## 4. No Monogram Letter Badges
- Do not generate arbitrary single-letter initial boxes (e.g., `[L]`, `[F]`, `[K]`, `[M]`) as placeholders for brands.
- Use clean brand typography, logos, or semantic category icons instead.

## 5. No Redundant Status Text Labels (Color & Placard State Sufficiency)
- **NEVER add explicit redundant text badges or labels** (e.g. `"Processing"`, `"Revalidate"`, `"Pending"`, `"Awaiting"`, `"Revalidation Completed"`) onto cards or placards when the placard's visual state (such as red placard tint, rose border/left stripe for pending/processing, cyan for revalidation pending, lime/mint lighter green for revalidation completed, or emerald for arrived/active) already serves as the definitive semantic signal.
- **Card Title / Line 1 Rule**: The title row must display **Brand name ONLY** (+ linked Card tag when tagged). Never clutter card headers with redundant state labels.
- **Definitive 4-Tier Visual Palette**:
  1. **Delivered / Arrived (Emerald Green)**: `card-arrived` (Background: `bg-[#062414]/75` [Dark] / `#EBF7EE` [Light], Border: `emerald-500/30`, Left Stripe: `emerald-500`). Code delivered and ready to use/approve.
  2. **Revalidation Completed (Lighter Shade of Green / Spring Lime)**: `card-revalidated` (Background: `bg-[#18522a]/90` [Dark] / `#EFFDE6` [Light], Border: `lime-400/45`, Left Stripe: `lime-400`). Visibly distinct lighter spring green; payment revalidation completed, awaiting coupon code delivery.
  3. **Pending Revalidation (Cyan)**: `card-revalidation` (Border: `cyan-500/40`, Left Stripe: `cyan-500`). Validation link active; OTP payment required.
  4. **Initial Pending / Processing (Rose / Red)**: `card-pending` (Border: `rose-500/30`, Left Stripe: `rose-500`). Order acknowledged; awaiting delivery or revalidation email.
- Color, borders, and left accent stripes convey lifecycle state cleanly with zero visual noise.
- Descriptive status text is strictly reserved for the detailed inspection modal / sheet (`(i)` detail sheet), keeping primary list feeds sleek, clean, and scannable.

---

## 6. Safeguards Audit Across Operations (Mandatory Pre-Flight Gate)
Before adding or modifying **any** operation, mutation, or workflow in the codebase (frontend or Google Apps Script backend), you must audit the operation against this checklist and integrate every safeguard that applies.

### Core Safeguards Matrix
Every operation must be evaluated against these 7 core safeguards:

1. **Universal Mutation Journaling (`UniversalMutationJournal`)**:
   - Record an optimistic mutation entry with a unique primary key and TTL (60s default).
   - Ensure in-flight items (`CARD_ADD`, `CARD_REPLACE`, `VOUCHER_APPROVE`, `VOUCHER_UNDO`, `CARD_TAG`, etc.) are merged and preserved during background syncs (`syncWithGoogleSheets`).
   - Clean up journal entries when remote confirmation is received.

2. **Ephemeral Quarantine & Anti-Collision Shield**:
   - For destructive/archival actions (block, delete, dismiss), quarantine the identifier to prevent ghost re-appearance during background polling.
   - For constructive/replacement/activation actions (add, replace, undo, reactivate), explicitly call `UniversalMutationJournal.unquarantine(id)` so valid entities are never suppressed by previous tombstones.

3. **Zero-Delay Network Execution & 3-Second Interactive Undo Banner**:
   - Destructive or high-impact user actions (Redeem, Clear, Block, Block & Replace, Mark as Sold, Dismiss, Add Misc) dispatch their backend mutation immediately with zero delay.
   - Concurrently, a 3-second interactive floating Undo banner (`showCardUndoBanner` / `showUndoToast` / `showClearUndoToast`) is presented to the user.
   - Clicking Undo immediately rolls back local memory state, unquarantines identifiers, updates the UI, and dispatches an explicit compensating reverse mutation (`UNDO_REDEEM`, `UNBLOCK_CARD`, `UNDO_BLOCK_AND_REPLACE_CARD`, `UNMARK_VOUCHER_SALE`, `RESTORE_CLEARED_ACTIVITY`, `DELETE_MISC_VOUCHER`) to Google Sheets.

4. **Network & Server Failure Rollback Handler**:
   - Every `apiFetch(action, payload, { rollback: () => { ... } })` invocation must supply an explicit `rollback` callback.
   - If the network drops, request times out (35s), lock contention occurs, or authentication fails, the callback must cleanly revert local state to its pre-mutation snapshot and alert the user.

5. **Mutation Mutex Shield (`pendingMutationsCount`)**:
   - `apiFetch` must increment `pendingMutationsCount` on dispatch and decrement in `finally`.
   - Prevent background read syncs from firing during active write mutations; queue them safely via `syncPendingAfterMutation = true`.

6. **Backend Transactional Integrity & Idempotency (`CronCode.gs`)**:
   - Guard all mutations with `LockService.getScriptLock()`.
   - Check for pre-existing records: update in-place if an entry already exists instead of appending duplicate rows.
   - For dual-sheet transfers (`Active Cards` <-> `Blocked Cards`, `New` -> Brand Tab, Brand Tab -> `Redeemed`):
     - Verify source row existence before deletion.
     - Purge conflicting states from destination/source tabs (e.g., purge from `Blocked Cards` if activating as replacement).
     - Preserve all metadata and benefit column values.

7. **Filter & View Context Preservation**:
   - Ensure newly created or replaced entities inherit the user's active filter context (Bank, Category, Linked Account).
   - Never allow a newly replaced or added card to default to `NA` or another account if the user is scoped to a specific account filter.

### Operations Safeguard Registry

#### Operation: `ADD_MISC_VOUCHER` (Adding Custom Voucher to Misc Tab)
- **Safeguard 1 (Mutation Journal)**: Record optimistic entry `MISC_VOUCHER_ADD` with unique key `misc_${voucherCode}` and 60s TTL so the newly added voucher survives background syncs until backend flushes.
- **Safeguard 2 (Quarantine Shield)**: Call `UniversalMutationJournal.unquarantine(voucherCode)` to guarantee the new voucher code is never suppressed by prior stale tombstones.
- **Safeguard 3 (Optimistic Feedback & Undo)**: Optimistically prepend the voucher card to the active Misc view and dispatch `ADD_MISC_VOUCHER` immediately (Zero Delay). Display a 3-second interactive Undo banner (`"Added [Brand] voucher. Tap to Undo"`); tapping Undo immediately rolls back local state and sends `DELETE_MISC_VOUCHER` to remove the voucher from the sheet.
- **Safeguard 4 (Rollback Handler)**: Supply an explicit rollback handler to `apiFetch` that restores the pre-mutation voucher state, clears the journal entry, and displays an alert if the network drops or backend fails.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring conflicting background syncs.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guard backend handler `ADD_MISC_VOUCHER` with top-level `LockService.getScriptLock().tryLock(30000)`.
  - Automatically create and format the `Misc` sheet if missing, with `UNIFIED_BRAND_HEADERS` and frozen top row.
  - Scan existing rows for duplicate voucher code. If already present, update in-place idempotently; otherwise append row.
- **Safeguard 7 (Context Preservation)**: Retains active view on `Misc` carousel tab without jarring resets.
- **Mandatory Verification**: Tested with a temporary dummy row (`MISC-TEST-DUMMY`) pushed to the spreadsheet, verified live, and safely purged.

#### Operation: `BLOCK_CARD` (Card Archival)
- **Safeguard 1 (Mutation Journal)**: Record optimistic entry `CARD_BLOCK` with primary key `cardNum` and card details in `UniversalMutationJournal`.
- **Safeguard 2 (Quarantine Shield)**: Tombstone `cardNum` in ephemeral quarantine (`isCardQuarantined`) so background sync does not resurface the blocked card before backend commit.
- **Safeguard 3 (Zero-Delay Execution & 3s Interactive Undo Banner)**: Optimistically archive card in memory and dispatch `BLOCK_CARD` immediately to Google Sheets. Present a 3-second floating Undo banner (`"Blocked [Card] (••[Last4]). Tap to Undo"`). Tapping Undo restores memory state, unquarantines the card, and dispatches explicit reverse call `UNBLOCK_CARD` to move the card back to Active Cards.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` closure supplied to `apiFetch` that restores `previousCards`, resets filters, unquarantines `cardNum`, and saves local encrypted storage upon network or server error.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring conflicting syncs.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guarded by `LockService.getScriptLock()`.
  - Pre-purge any stale duplicates of `cardNum` from `Blocked Cards`.
  - Locate row in `Active Cards`, copy row with `BLOCKED` status and timestamp to `Blocked Cards`, and safely delete row from `Active Cards`.
- **Safeguard 7 (Context Preservation)**: If the active secondary account filter pill was locked to the blocked card's digits (`cardNum.slice(-4)`), check for remaining cards; if none match, smoothly relax filter to `'All'` to avoid presenting an empty card manager screen.
- **Mandatory Verification**: Tested via `VERIFY_SAFEGUARD_CARD_DUMMY` on live sheet and safely purged.

#### Operation: `BLOCK_AND_REPLACE_CARD` (Compound Atomic Swap)
- **Safeguard 1 (Mutation Journal)**: Record `CARD_REPLACE` with primary key `newCardNum` and secondary key `oldCardNum` in `UniversalMutationJournal`.
- **Safeguard 2 (Quarantine Shield)**: Tombstone `oldCardNum`; explicitly call `UniversalMutationJournal.unquarantine(newCardNum)` to prevent replacement card collision with any prior tombstones.
- **Safeguard 3 (Zero-Delay Execution & 3s Interactive Undo Banner)**: Optimistically swap the card in memory, re-render the list immediately, and dispatch `BLOCK_AND_REPLACE_CARD` to backend immediately. Present a 3-second interactive Undo banner (`"Replaced with ••[Last4]! Tap to Undo"`). Tapping Undo rolls back local memory and dispatches `UNDO_BLOCK_AND_REPLACE_CARD` to Google Sheets.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` closure in `apiFetch` that restores `previousCards`, restores `currentCardAccountFilter`, unquarantines `oldCardNum`, removes `newCardNum` from the journal, and refreshes the view.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, blocking concurrent read collisions.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guarded by `LockService.getScriptLock()`.
  - Dual-sheet purge: remove `newCardNum` and `oldCardNum` from `Blocked Cards` to prevent duplicate or conflicting status.
  - Archive `oldCardNum` to `Blocked Cards` with `Replaced_By_Card` metadata.
  - Check `Active Cards`: if `newCardNum` is already present, update in-place; otherwise append new row with eligible RuPay benefits and linked account.
- **Safeguard 7 (Context Preservation)**: If the active view was scoped to the old card's account or last-4 digits, smoothly migrate `currentCardAccountFilter` to the replacement card's last-4 or linked account digits so the replacement card remains immediately visible.
- **Mandatory Verification**: Tested live on Google Sheets via `VERIFY_SAFEGUARD_CARD_DUMMY` and purged with 0 residual artifacts.

#### Operation: `ADD_CARD` (Card Portfolio Onboarding)
- **Safeguard 1 (Mutation Journal)**: Record optimistic entry `CARD_ADD` with primary key `cardNumber` in `UniversalMutationJournal` with 60s TTL.
- **Safeguard 2 (Quarantine Shield)**: Explicitly unquarantine `cardNumber` on addition.
- **Safeguard 3 (Optimistic UI & Feedback)**: Prepend card to active list, close modal, update count badges, and toast immediately.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` callback supplied to `apiFetch('ADD_CARD', ...)` that reverts `state.cards` to snapshot and removes journal entry upon error.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount`.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guarded by `LockService.getScriptLock()`.
  - Scan `Active Cards` for pre-existing `cardNumber`. If found, update in-place; otherwise append new row.
- **Safeguard 7 (Context Preservation)**: Retains active bank and category filters; pre-fills modal with active filter values.

#### Operation: `TAG_CARD` (Metadata Annotation on Vouchers)
- **Safeguard 1 (Mutation Journal)**: Record optimistic tag mutation entries for all selected voucher codes via `MutationJournal.recordTags`.
- **Safeguard 2 (Quarantine Shield)**: N/A (metadata update only).
- **Safeguard 3 (Optimistic UI)**: Trigger tactile slide animation, update local card tags across `state.vouchers` and `state.staging`, and collapse dock.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` callback supplied to `apiFetch('TAG_CARD', ...)` reverting vouchers and staging to previous snapshot and purging journal tags on server error.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, deferring background syncs during write.
- **Safeguard 6 (Backend Transactional Integrity)**: In-place cell update in brand and staging sheets under script lock.
- **Safeguard 7 (Context Preservation)**: Preserves active brand tab and selection state.

#### Operation: `UNDO_REDEEM` (Restorative Reverse Transfer)
- **Safeguard 1 (Mutation Journal)**: Record optimistic restoration entry in `UniversalMutationJournal`.
- **Safeguard 2 (Quarantine Shield)**: Call `UniversalMutationJournal.unquarantine(codes)` so the restored voucher is immediately visible.
- **Safeguard 3 (Optimistic UI)**: Move voucher from Redeemed archive back to active brand tab immediately with success toast.
- **Safeguard 4 (Rollback Handler)**: Reverts local list and journal entries if backend return indicates error.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount`.
- **Safeguard 6 (Backend Transactional Integrity)**: Locate row in `Redeemed`, copy back to original brand tab with `ACTIVE` status, and safely delete row in descending index order under script lock. Self-healing check verifies if row was already restored.
- **Safeguard 7 (Context Preservation)**: Restores card into matching brand carousel view.

#### Operation: `MARK_REVALIDATED` & `UNDO_MARK_REVALIDATED` (Transitioning Staged Voucher to Revalidated / Awaiting Code)
- **Safeguard 1 (Mutation Journal)**: Dedicated persistent revalidation register (`rupay_revalidated_bids`) with zero time-based expiration. Remains active until the real delivered coupon code arrives. In addition, dual-recorded in `UniversalMutationJournal` for short-term in-flight sync safety.
- **Safeguard 2 (Quarantine Shield)**: N/A (constructive in-place state transition). Explicitly excluded from destructive quarantine so it is never suppressed.
- **Safeguard 3 (Zero-Delay Execution & 3s Interactive Undo Banner)**: Optimistically update voucher to lighter shade of green `card-revalidated` status ("Revalidation Completed" in detail sheet), clear payment link button, close detail sheet, and present 3-second interactive Undo banner (`"Revalidation completed ([Brand]). Tap to Undo"`). Tapping Undo reverts state to `PENDING_REVALIDATION` and dispatches compensating `UNDO_MARK_REVALIDATED` to Google Apps Script.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` callback supplied to `apiFetch('MARK_REVALIDATED', ...)` reverting `state.staging` to snapshot and removing the booking ID from the revalidated register upon network or server error.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring conflicting background syncs.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guarded by top-level `LockService.getScriptLock().tryLock(30000)`.
  - In `MARK_REVALIDATED`: updates `New` tab (Col 2: `'Pending'`, Col 8: `'REVALIDATED'`, Col 12: `''`).
  - In `UNDO_MARK_REVALIDATED`: updates `New` tab (Col 2: `'Pending revalidation'`, Col 8: `'PENDING_REVALIDATION'`, Col 12: `payUrl`) and removes booking ID from `REVALIDATED_BOOKINGS`.
  - Stores revalidated booking ID in `PropertiesService.getScriptProperties()` under `REVALIDATED_BOOKINGS`.
  - In `recordPendingRevalidation()`: check `isBookingRevalidated(targetBid)` before modifying rows; never clobbers a user-revalidated voucher back to `PENDING_REVALIDATION`, while cleanly transitioning fresh Stage 1 vouchers (`PENDING`/`PROCESSING`) to `PENDING_REVALIDATION` with their payment URL.
  - **Timestamp Invariant**: When a pending revalidation email or link arrives, Col 9 (`Timestamp`) is strictly immutable and preserved as the initial ingested time. Thread initial date is anchored on ingestion, and `recordPendingRedemption` heals Col 9 to the earlier Stage 1 acknowledgment time if the revalidation email was ingested first.
  - When real coupon code arrives in `processRupayEmails`: purge booking ID from `REVALIDATED_BOOKINGS`.
  - Removed legacy hardcoded auto-heal clobber list.
- **Safeguard 7 (Context Preservation)**: Retains active filter and selection context.
- **Mandatory Verification**: Tested via `VERIFY_SAFEGUARD_REVAL_DUMMY` on live sheet and safely purged, asserting that status transitions to `REVALIDATED` and timestamp remains unchanged across the entire revalidation transition and repeat crons.

#### Operation: `RECORD_INCOME_ENTRY` (Recording Sales into "RuPay Income" and Charges into "RuPay Charges")
- **Safeguard 1 (Mutation Journal)**: Record optimistic transaction entry in `state.income.ledger` with unique ID (`inc-${Date.now()}`), preventing loss during background syncs.
- **Safeguard 2 (Quarantine Shield)**: N/A (constructive ledger append). Explicitly listed in `nonQuarantineActions` so newly recorded entries are never suppressed by tombstone filters.
- **Safeguard 3 (Optimistic UI & Feedback)**: Optimistically insert into `state.income.ledger` in date-sorted order (newest first in UI feed), recalculate KPIs (Gross, Charges, Net), close modal, and display instant toast (`Logged [Title] sale (+Rs.X) in RuPay Income!` or `Deducted [Title] (-Rs.X) in RuPay Charges!`).
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` closure supplied to `apiFetch('RECORD_INCOME_ENTRY', ...)` that filters out `newEntry.id` from `state.income.ledger`, re-renders the tracker, and alerts the user if the server or network fails.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring conflicting background reads until write completes.
- **Safeguard 6 (Backend Transactional Integrity & Date-Sorted Insertion)**:
  - Guarded by top-level `LockService.getScriptLock().tryLock(30000)`.
  - Routes sales (`price > 0`) to `"RuPay Income"` and card charges (`price < 0`) to `"RuPay Charges"`. Both tabs are registered under `SYSTEM_SHEET_NAMES`.
  - Automatically initializes `"RuPay Charges"` tab with bold headers and frozen top row if missing.
  - Scans target sheet's date column and inserts new row at the exact chronological ascending position based on transaction date (before the first row with a later date, and before `TOTAL` row).
  - Maintains `TOTAL` summary row dynamically on every write.
- **Safeguard 7 (Context Preservation)**: Preserves active year filter (`all`, `2026`, `2025`, `2024`), type filter (`all`, `income`, `charge`), and active search query.
- **Mandatory Verification**: Tested live on Google Sheets via `VERIFY_SAFEGUARD_INCOME_DUMMY` across both sheets with sorted insertion verified and 0 residual artifacts.

#### Operation: `MARK_VOUCHER_AS_SOLD` (Direct Ingested/Brand -> Redeemed Tab + RuPay Income Entry for Cult & Amazon Prime)
- **Safeguard 1 (Mutation Journal)**: Optimistic entries recorded in `UniversalMutationJournal` for voucher redemption and income ledger creation, protecting state against background polling.
- **Safeguard 2 (Quarantine Shield)**: Booking ID and code are quarantined to prevent ghost re-emergence in Ingested or Brand feeds during syncs.
- **Safeguard 3 (Optimistic UI & Feedback)**: Voucher immediately leaves staging/brand views and unshifts into `redeemedList`; sale unshifts into `state.income.ledger` with dynamic KPI recalculations; dispatches `MARK_VOUCHER_AS_SOLD` to backend immediately with zero delay; displays 3-second interactive Undo banner (`"Marked [Brand] sold to [Buyer] (+₹[Price]) · Tap to Undo"`). Tapping Undo reverts local state and sends `UNMARK_VOUCHER_SALE` to backend.
- **Safeguard 4 (Rollback Handler)**: Explicit `rollback` callback supplied to `apiFetch('MARK_VOUCHER_AS_SOLD', ...)` restoring `state.staging` or `state.vouchers`, un-redeeming from `redeemedList`, removing the ledger sale, and alerting user upon network failure.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring background syncs until atomic write completes.
- **Safeguard 6 (Backend Transactional Integrity & Single-Lock Atomicity)**:
  - Guarded by `LockService.getScriptLock().tryLock(30000)`.
  - Atomically moves row to `"Redeemed"` sheet (from `"New"` or brand tab) and inserts sorted row into `"RuPay Income"` sheet with updated `TOTAL` summary row.
- **Safeguard 7 (Context Preservation)**: Preserves active brand filter and closes inspection bottom sheets.
- **Mandatory Verification**: Tested live on Google Sheets via `VERIFY_SAFEGUARD_SOLD_DUMMY` asserting row movement to Redeemed, sorted insertion into RuPay Income, and 0 residual artifacts.

### Mandatory Verification via Dummy Row Testing
Whenever adding a new operation or feature change:
1. **Insert a tagged dummy row** into the relevant Google Sheet tab (e.g. `TEST-DUMMY-SAFEGUARD` / card `6521999900000001`).
2. **Verify live behavior**:
   - Optimistic UI update and Undo toast appearance.
   - Backend transactional execution and idempotency against repeat calls.
   - Rollback handling under forced error/network interruption.
3. **Post-test cleanup**: Automatically delete the dummy row from the spreadsheet and verify that zero residual test artifacts remain.
