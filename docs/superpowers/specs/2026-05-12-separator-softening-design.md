# Design — Separator softening (Stream C / #9)

> Status: approved 2026-05-12. Implementation plan to follow via `writing-plans`.
> Scope: SCSS-only visual polish across four web parts. No behaviour change, no new property-pane config.

## Problem

Result rows — and several other list/section surfaces — draw visible `1px` separator lines (`neutralLighter #f3f2f1` between rows; `neutralLight #edebe9` for structural dividers; one `2px` header underline). On a white SharePoint page these read as heavier than wanted. There's already a row-hover affordance (`.resultCard:hover { background: neutralLighterAlt #faf9f8 }`) — almost invisible — so the visual weight is doing double duty with the hover. The user wants the lines softened and the hover to carry the "this is a row you can interact with" signal.

## Decision

Two rules, applied consistently across **Search Results** (List + Compact layouts), **Search Manager / Admin Manager**, the **result Detail panel**, and the **Filters drawer**:

### Rule A — list-row separators → removed; hover does the work

Every place that today draws a `border-bottom` *between rows in a list* drops the border (rows are separated by their existing padding) and gets a stronger row hover/focus treatment, defined once in a shared mixin:

```scss
// src/styles/rowSurface.scss
@import 'motion';

// Standard treatment for an interactive row in a list (result cards, manager
// list rows, filter-option rows, …). No border at rest — separation is the
// row's own padding; hover/focus carries the affordance.
@mixin interactive-row {
  @include motion(background-color box-shadow, 150ms, ease);

  &:hover,
  &:focus-within {                                              // keyboard users get the same highlight when tabbing
    background-color: var(--neutralLighter, #f3f2f1);           // clear but still light — was #faf9f8 (barely visible)
    box-shadow: inset 3px 0 0 0 var(--themePrimary, #0078d4);   // left accent via inset shadow → no layout shift vs a real left border
  }
}
```

Surfaces that adopt `@include interactive-row` (and drop their `border-bottom`):

| File | Class(es) | Today |
|------|-----------|-------|
| `spSearchResults/components/SpSearchResults.module.scss` | `.resultCard` (List), `.compactRow` (Compact) | `border-bottom: 1px solid neutralLighter`; weak `:hover` tint |
| `spSearchManager/components/SpSearchManager.module.scss` | the saved-search / history / collection list rows (the `neutralLighter`-bordered, `:last-child border-bottom: none` rows around lines 204, 343, 428, 1186, 1299) | same pattern |
| `spSearchFilters/components/SpSearchFilters.module.scss` | the filter-option list rows (`neutralLighter`-bordered row, ~line 797) | same pattern |
| `spSearchResults/components/SpSearchResults.module.scss` | detail-panel related-documents / version-list rows (the `.detailPanel*`-prefixed classes that draw a `border-bottom` between repeated items) | same pattern |

> Implementation enumerates the exact class names per file by grepping each `*.module.scss` for `border-bottom: 1px solid "[theme:neutralLighter…]"` on a row class — the rule is "if it's a separator between items in a list, it gets `@include interactive-row` and loses its border-bottom".
>
> **Rule A vs Rule B, if a given border is ambiguous:** does it sit between two instances of the *same repeated element* (cards, list items)? → Rule A (remove + `@include interactive-row`). Does it sit between *distinct sections / headers / toolbars*? → Rule B (soften, keep).

### Rule B — structural dividers → softened, not removed

Lines that *delimit sections* rather than rows — web-part container border, panel/section headers, toolbar bottom borders, the title hover-card divider (`.hoverCardDivider`), detail-panel section dividers, filter-group header borders, and the lone `2px` header underline — stay (no hover target; removing them would blur section boundaries) but go one step lighter:

- `border: 1px solid "[theme:neutralLight, default: #edebe9]"` → `border: 1px solid "[theme:neutralLighter, default: #f3f2f1]"` (and the `var(--neutralLight, #edebe9)` fallback line alongside it → `var(--neutralLighter, #f3f2f1)`)
- the single `2px solid neutralLight` header underline (`SpSearchResults.module.scss` ~line 941) → `1px solid neutralLighter`

### Not touched

- `.metaSeparator` — the tiny `3px` `neutralTertiaryAlt #c8c6c4` bullet between metadata items (a dot, not a line).
- Vertical-tab active underline (`themePrimary` `border-bottom-color`) — an intentional selected-state indicator, not a separator.
- `DebugPanel.module.scss` dark borders (`#333` / `#444` / `#2a2a2a`) — intentionally dark; the Debug Panel is a dark-themed dev surface.
- `border-left: 3px solid themePrimary` on `.promotedCard` / `border-left: 4px solid themePrimary` on the Manager selected row — existing intentional accents; leave as-is (they don't conflict with the new hover inset-shadow accent since they're not the same elements, but if a row both is "selected" and "hovered", the inset shadow + the real left border stack harmlessly).

## Plumbing

New shared partial `src/styles/rowSurface.scss` exporting `@mixin interactive-row` (uses `@include motion(...)` from the existing `src/styles/motion.scss`, so the transition is `prefers-reduced-motion`-aware per Found.D6). `@import 'rowSurface';` into each of the three web-part `*.module.scss` files that adopt Rule A. Rule B is per-declaration edits, no mixin.

## Verification

- `sass` typings regenerate cleanly (heft `build:sass`), SCSS compiles.
- `npm run type-check` clean (no TS impact, but the build pulls the SCSS).
- `npm test` green — including the axe a11y smoke (`tests/a11y/smokeAxe.test.tsx`); the `:focus-within` clause is the relevant a11y bit, and axe won't flag a removed cosmetic border.
- `npm run package` + `npm run check:bundles` green (CSS delta is a few hundred bytes; well within budget).
- No unit test — pure colour/border. (A visual-regression harness is out of scope.)

## Out of scope / follow-ups

- Stream C #7 (result-item click behaviour) and #8 (image preview) — separate designs, after this.
- Any broader Fluent-theme token rationalisation (e.g. introducing project CSS custom properties for separator colour) — not now.
