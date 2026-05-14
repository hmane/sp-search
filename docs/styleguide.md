# SP Search — Style Guide (T1.D12)

> Owned by T1 / Modern UI Quality. Single source of truth for the
> shared visual + interaction tokens used across every shipped web
> part. The "visual regression suite" half of T1.D12 (CI-side
> screenshot diff) is deferred — this doc captures the durable
> contract that future regression coverage validates against.

## Breakpoints

Two shared breakpoints live in `src/styles/breakpoints.scss` and are
the ONLY mobile / tablet pivot points any module SCSS should consume.
Raw `@media (max-width: Npx)` queries against the documented
breakpoint values are also allowed (Results uses raw queries for the
sub-400 px tier).

| Mixin | Range | Use case |
|-------|-------|----------|
| `@include mobile-and-below` | 0–639 px | Phones, narrow web part columns |
| `@include tablet-and-below` | 0–1023 px | Phones + tablets / split-screen |
| `@include desktop-and-above` | 1024 px+ | Wide layouts |

Sub-tier: a raw `@media (max-width: 399px)` is allowed for very-narrow
phone targets (single-column grids, condensed pill bars). See
`SpSearchResults.module.scss` and `SpSearchFilters.module.scss` for
the canonical patterns.

**Don't** introduce new breakpoints. The audit explicitly closes the
prior drift across 480 / 640 / 768 / 900 / 1023 px values.

## Theme tokens (colour)

SP Search inherits the SharePoint Modern host theme — see
[`docs/theming.md`](./theming.md) for the full token list. The most-
used token families per surface family:

| Surface | Tokens |
|---------|--------|
| Page chrome | `white`, `neutralLighterAlt`, `neutralLighter` |
| Body text | `bodyText`, `bodySubtext`, `neutralSecondary` |
| Borders | `neutralLighter`, `neutralTertiary` |
| Accent | `themePrimary`, `themeDarkAlt`, `themeLighterAlt` |
| Status | `errorBackground`, `successBackground`, `severeWarningBackground` |
| Disabled | `neutralQuaternaryAlt`, `neutralTertiaryAlt` |

Every `.module.scss` uses both the SPFx `[theme:token, default:#hex]`
form (parsed by `sp-css-loader`) AND the standard
`var(--token, #hex)` form (parsed by browsers when Fluent UI's
ThemeProvider injects the CSS variable). The double form is the
contract every theme-token consumer must honour.

## Spacing

The shared SCSS modules use a 4 px grid:

- `4 / 8 / 12 / 16 / 24 / 32 / 40 / 48 / 64` px multiples.
- Avoid odd values (5, 7, 13...) — they read as mistakes against the
  rest of the suite.
- Inline `style={{ padding: 12 }}` is acceptable for one-off Fluent-
  wrapped surfaces (Pivot tab spacing, Modal bodies) but prefer SCSS
  classes for repeated surface shapes.

## Type scale

Loose anchors used across the suite (overridable per surface via
theme tokens):

- Page section title: 18 px, weight 600
- List row title: 14-15 px, weight 500
- Body / meta line: 13 px, regular
- Caption / hint: 12 px, `neutralSecondary` colour
- Micro / keycap chip: 11 px, monospace (Consolas)

iOS Safari auto-zooms inputs whose font-size is below 16 px on
focus. On mobile (`@include mobile-and-below`), every search-input
and KQL-textarea selector must declare `font-size: 16px`. See T1.D10
implementation in `SpSearchBox.module.scss`.

## Motion

Default motion rule: 150 ms `ease` on hover/focus background
transitions. Anything longer feels laggy on the SharePoint host.

`prefers-reduced-motion: reduce` is honoured via the universal-
selector media query at the bottom of every `.module.scss`. The
shared `src/styles/motion.scss` mixin
`@include motion-safe { ... }` wraps surface-specific animations
that should disable when the user has opted out. See
[`docs/theming.md#motion-reduction-prefers-reduced-motion`](./theming.md#motion-reduction-prefers-reduced-motion).

## Empty states

Per T1.D5: the **idle / pre-search** empty state uses the `Search`
icon; **post-search zero-result** uses `SearchAndApps`. **Never**
use `SearchIssue` — it reads as a warning triangle and
miscommunicates that something went wrong. Recovery buttons
("Clear all filters" / "Start over") sit below the icon + copy.

## Loading

One shimmer idiom per surface, matching the post-load shape:

- **Results list / Compact / Card / People / Gallery**: skeleton row
  layouts (shape-matched).
- **Manager panel**: header line + 3 list-row skeletons (T1.D3).
- **Filters panel**: 3 filter-group skeleton blocks.
- **Detail panel preview**: full-pane `ShimmerElementType.line` block.

Don't introduce per-component spinners for layout-bearing surfaces —
the audit closed three competing loading idioms in T1.D3.

## Buttons and actions

- **Primary action**: Fluent `<PrimaryButton>` or `<DefaultButton
  primary>`. One per dialog.
- **Secondary actions**: `<DefaultButton>` or `<IconButton>` with
  tooltip.
- **Destructive actions**: `<PrimaryButton>` with red theme (e.g.
  Force dispose in DebugPanel) AND a `window.confirm`. Never a bare
  destructive primary without confirmation.

Touch targets: ≥44 px on mobile per iOS HIG / Android Material. The
Filters drawer toggle ships with `min-height: 44px` for that
reason; new mobile action buttons should match.

## Tooltips

Use Fluent `<TooltipHost>` everywhere — never the HTML `title=`
attribute on its own. TooltipHost adds the aria-described-by
relationship a screen reader needs. The browser-native `title=` can
ride along on the same element for hover-while-loading parity but
isn't a replacement.

Disabled-button tooltips MUST switch copy between enabled and
disabled states (T2.D10 pattern):

- Enabled: "Save the current query and filters"
- Disabled: "Type a query or apply a filter to enable Save"

Static "this button does X" copy on a greyed button reads as a
riddle.

## Accessibility floor

- All custom buttons have `aria-label` matching the visual copy.
- Disabled state is announced via `disabled={true}` + the
  context-aware tooltip above, plus `aria-label` updates for SR
  users.
- Panels and modals use Fluent's built-in FocusTrapZone (don't roll
  your own).
- Skip-link / focus management: see [`docs/accessibility.md`](./accessibility.md).

## Visual regression suite (deferred)

The CI-side screenshot-diff pipeline is deferred to the Azure
DevOps pipeline work outside this repo. The acceptance signal
"visual regression runs on every PR and fails on uncalibrated diffs"
remains open as a Sprint 6+ pipeline deliverable.

Until that lands, the manual smoke-checklist (see
[`docs/release-smoke-checklist.md`](./release-smoke-checklist.md))
covers the screenshot capture cadence for each release.
