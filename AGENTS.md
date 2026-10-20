# RuPay Vault — Developer Guidelines & Memory Rules

## Critical UI & Interaction Rules

### 1. Mobile Pass Cards (STRICT)
- **NEVER add individual action buttons ("Share via WhatsApp", "Redeem", etc.) to mobile pass cards.**
- Mobile pass cards in the active vault and Activity Log must NEVER have individual action buttons.
- All voucher actions (WhatsApp sharing, redemption) MUST be executed exclusively through the bottom **Floating Action Dock (Interlock Master Combo FAB)** via Select Mode.
- Mobile pass cards must remain clean and uncluttered:
  1. **Top Row**: Selection checkbox (visible ONLY when Select Mode is active), Brand tag / Pass #, and status badge.
  2. **Timestamp Row**: Ingestion or validity timestamp.
  3. **Embossed Code Box**: Monospace Code & PIN with Copy button.
- Tapping anywhere on an active mobile card toggles selection.

### 2. Floating Action Dock & Redemption Workflow
- **Resting State**: 56px unified Interlock Master Combo FAB (equal-split WhatsApp + Double Tick).
- **Expanded State**: Two 56px stacked circles (WhatsApp green circle on top, Double Tick emerald circle below) without text labels.
- **Badge Visibility**: The selection count badge appears ONLY on the resting master bubble. When expanded into stacked circles and the cross button, the count badge disappears completely. After clicking the cross button, the count badge reappears on the master bubble.
- **Exit Animation**: 700ms pure horizontal slide-through (`anim-swipe-horizontal-700` + `anim-collapse-700`) with 0° tilt/wobble.
- **Select Mode Reset**: On redemption, the app immediately exits Select Mode and resets the dock to the resting Interlock FAB.
- **Undo Toast**: Exactly 3.0-second countdown with reverse horizontal glide rollback (`anim-rollback-700`).

### 3. Navigation & Default State
- The app must **always default to ⚡ Activity Log on page load/reload**.
- Deck selector must NOT show false-active glowing borders on unselected tabs.

---

## Strict Backend & Data Integrity Rules

### 1. Booking ID is the Sole Canonical Identifier for Idempotency (STRICT)
- **All idempotency and deduplication checks across Google Sheets tabs MUST be performed EXCLUSIVELY on Booking ID (`RUPS-...`).**
- **NEVER use dual-identifier matching (matching on Voucher Code) for idempotency.**
- **Reasoning**:
  1. Voucher codes are absent or placeholder (`PENDING`) during initial acknowledgement and booking emails.
  2. Voucher codes may carry markdown artifacts, prefix quotes (`'`), or formatting differences across stages.
  3. The Booking ID (`RUPS-...`) is generated directly by RuPay/Truztee and is the permanent, unique primary key across the entire voucher lifecycle.
- **Enforcement Scope**:
  - `APPROVE_STAGING`: Inspect destination Brand tab strictly by Booking ID before appending. If found, mark `ACTIVE` and do not append a duplicate row.
  - `recordPendingRedemption`: Inspect `Redeemed`, `"New"`, and Brand tabs strictly by Booking ID.
  - `fulfillVoucherCode`: Match and heal rows strictly by Booking ID.
  - Any future migration, reconciliation, or movement scripts must adhere to Booking ID-only matching for idempotency.

### 2. Strict Voucher Lifecycle Transition Invariant
- Canonical flow: `"New"` (Staging) $\rightarrow$ Brand Tab (Active Vault) $\rightarrow$ `"Redeemed"` (Archive).
- **Direct `"New"` $\rightarrow$ `"Redeemed"` transition is STRICTLY FORBIDDEN** on both backend and frontend. Vouchers in staging MUST be approved to a brand tab first.
- `"New"` tab is a system staging buffer and must always remain in `systemTabs` to isolate it from redemption scans.
