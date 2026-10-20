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

### 1. Hierarchical Idempotency Model (Primary RUPS Booking ID, Fallback Voucher Code) (STRICT)
- **Email Ingestion into "New" Tab (Strict RUPS Invariant)**:
  - Email ingestion (`recordPendingRedemption`, `fulfillVoucherCode`) MUST be keyed strictly by RuPay/Truztee Booking ID (`RUPS-...`).
  - Voucher codes in `"New"` can be absent or placeholder (`PENDING`) during initial acknowledgement, so code-based ingestion matching is strictly forbidden to prevent collision.
- **Brand Tab Mutations & Lifecycle Operations (Hierarchical Invariant)**:
  - Canonical Brand Tabs contain two classes of vouchers:
    1. Standard RuPay Select vouchers (Apollo, Lakme, Kalyan, etc.) carrying a `RUPS-...` Booking ID in Column B.
    2. Direct partner membership passes (Times Prime Ace, Decathlon, Swiggy One direct, legacy passes) which do NOT have a `RUPS-...` code (Column B contains partner pass names or is blank), but have a globally unique voucher code in Column C.
  - Therefore, all brand tab mutations (`APPROVE_STAGING`, `MOVE_TO_REDEEMED`, `UNDO_REDEEM`, `findRowInBrandTabs`):
    - **Primary Identifier**: If Booking ID is present and starts with `RUPS-`, match strictly on `RUPS-...` Booking ID.
    - **Fallback Identifier**: If Booking ID does NOT start with `RUPS-` (or is absent/placeholder), match on the unique Voucher Code (ignoring `PENDING`, `-`, and blanks).
  - This hierarchical model prevents duplicate rows while ensuring non-RUPS partner vouchers can be moved to `"Redeemed"` and undone seamlessly without getting trapped or locked out.

### 2. Strict Voucher Lifecycle Transition Invariant
- Canonical flow: `"New"` (Staging) $\rightarrow$ Brand Tab (Active Vault) $\rightarrow$ `"Redeemed"` (Archive).
- **Direct `"New"` $\rightarrow$ `"Redeemed"` transition is STRICTLY FORBIDDEN** on both backend and frontend. Vouchers in staging MUST be approved to a brand tab first.
- `"New"` tab is a system staging buffer and must always remain in `systemTabs` to isolate it from redemption scans.

### 3. Unified Column Schema (No '#' Column) (STRICT)
- The Google Sheet does NOT use a `#` (row index) column in any tab.
- All brand tabs, `New`, `Redeemed`, and `Shared_Activity` start directly with their semantic primary keys:
  - **Brand tabs**: `Booking ID` (Col A), `Code` (Col B), `PIN` (Col C), `Value` (Col D), `Date of Issue` (Col E), `Account` (Col F), `Card` (Col G), `Status` (Col H), `Appended Timestamp` (Col I), `Display in UI` (Col J).
  - **New tab**: `Booking ID` (Col A), `Code` (Col B), `PIN` (Col C), `Value` (Col D), `Date of Issue` (Col E), `Account` (Col F), `Card` (Col G), `Status` (Col H), `Appended Timestamp` (Col I), `Display in UI` (Col J), `Target Brand` (Col K).
  - **Redeemed tab**: `Brand` (Col A), `Booking ID` (Col B), `Code` (Col C), `PIN` (Col D), `Value` (Col E), `Date of Issue` (Col F), `Account` (Col G), `Card` (Col H), `Status` (Col I), `Redeemed Timestamp` (Col J), `Original Tab` (Col K), `Appended Timestamp` (Col L), `Display in UI` (Col M).
  - **Shared_Activity tab**: `Code` (Col A), `Brand` (Col B), `Amount` (Col C), `PIN` (Col D), `Shared Timestamp` (Col E), `Status` (Col F), `Platform` (Col G), `Display in UI` (Col H).
- Renumbering loops (`setValues(nums)`) are strictly forbidden, as row numbering is handled natively by Google Sheets and card numbering is rendered dynamically by the UI.
