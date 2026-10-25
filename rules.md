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
