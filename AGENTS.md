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
