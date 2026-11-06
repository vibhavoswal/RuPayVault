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

---

## 5. Safeguards Audit Across Operations (Mandatory Pre-Flight Gate)
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

3. **Optimistic Grace Period & Interactive Undo Banner**:
   - Destructive or high-impact user actions (Redeem, Clear, Block, Block & Replace) must display a 3–5 second countdown floating Undo banner (`showCardUndoBanner` / `showUndoBanner`).
   - Clicking Undo must immediately cancel the scheduled backend dispatch, restore memory state, unquarantine identifiers, and update the UI.

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
- **Safeguard 3 (Optimistic Feedback & Undo)**: Optimistically prepend the voucher card to the active Misc view immediately. Display a 5-second interactive Undo banner (`"Added [Brand] voucher. Tap to Undo"`) allowing the user to abort before backend commit.
- **Safeguard 4 (Rollback Handler)**: Supply an explicit rollback handler to `apiFetch` that restores the pre-mutation voucher state, clears the journal entry, and displays an alert if the network drops or backend fails.
- **Safeguard 5 (Mutation Mutex)**: Executed via `apiFetch`, incrementing `pendingMutationsCount` and deferring conflicting background syncs.
- **Safeguard 6 (Backend Transactional Integrity & Idempotency)**:
  - Guard backend handler `ADD_MISC_VOUCHER` with `LockService.getScriptLock().waitLock(30000)`.
  - Automatically create and format the `Misc` sheet if missing, with `UNIFIED_BRAND_HEADERS` and frozen top row.
  - Scan existing rows for duplicate voucher code. If already present, update in-place idempotently; otherwise append row.
- **Safeguard 7 (Context Preservation)**: Retains active view on `Misc` carousel tab without jarring resets.
- **Mandatory Verification**: Tested with a temporary dummy row (`MISC-TEST-DUMMY`) pushed to the spreadsheet, verified live, and safely purged.

### Mandatory Verification via Dummy Row Testing
Whenever adding a new operation or feature change:
1. **Insert a tagged dummy row** into the relevant Google Sheet tab (e.g. `TEST-DUMMY-SAFEGUARD` / card `6521999900000001`).
2. **Verify live behavior**:
   - Optimistic UI update and Undo toast appearance.
   - Backend transactional execution and idempotency against repeat calls.
   - Rollback handling under forced error/network interruption.
3. **Post-test cleanup**: Automatically delete the dummy row from the spreadsheet and verify that zero residual test artifacts remain.
