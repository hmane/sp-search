# SP Search — Accessibility Conformance Statement (WCAG 2.1 AA)

> Owned by Foundations track (Found.D6). Scoped — covers axe-core-tested surfaces only. Manual screen-reader pass on the full surface is v1.1+ work.

## Scope

This statement covers the four surfaces exercised by `tests/a11y/smokeAxe.test.tsx` (axe-core CI gate) plus the two A11Y-004 / A11Y-005 closures landed in this release:

1. **Search Box mode toggle** (`SpSearchBox.tsx:735`) — semantic `<fieldset>` + visually-hidden `<legend>`, role="radio" buttons with `aria-checked`. (A11Y-004 closed.)
2. **Search Box scope selector** (`SpSearchBox.tsx:723`) — Fluent UI Dropdown with `aria-describedby` linking to a hidden description span. (A11Y-005 closed.)
3. **Empty state markup** (`SearchResults.tsx`) — `role="status" aria-live="polite"`.
4. **Detail panel close button** (`ResultDetailPanel.tsx`) — explicit `aria-label`.

## Conformance

We claim WCAG 2.1 Level AA conformance for the surfaces enumerated above. All other surfaces inherit a baseline level via the SPFx host and Fluent UI v8 (which itself conforms to WCAG 2.1 AA), but have not been independently audited as of v1.0.

## Testing approach

- **Static analysis (CI gate)** — `axe-core` via `jest-axe` on every PR. New violations on the four enumerated surfaces fail the build via `tests/a11y/smokeAxe.test.tsx`.
- **Motion preference** — `prefers-reduced-motion: reduce` honored across module.scss files via the universal-selector media query (and the shared `src/styles/motion.scss` mixin available for surface-specific application by T1.D9). Verified via `grep -rn "prefers-reduced-motion" src/styles src/webparts`.
- **Keyboard navigation** — Tab order verified manually for the Search Box, Filters drawer, Detail panel close button, and Manager tabs. Esc closes any open panel.
- **Focus visible** — relies on Fluent UI v8 default focus rings; no custom focus-ring suppression in `*.module.scss`.

## Known limitations (v1.0)

- No manual screen-reader (NVDA / JAWS / VoiceOver) pass on file. Full conformance verification is v1.1+.
- DataGrid layout (DevExtreme) accessibility relies on DevExtreme's own a11y posture; we do not re-audit it.
- Filters drawer does not yet ship `FocusTrapZone` (T1.D1 dep — Sprint 5).
- ManageAccess panel does not yet exist (T2.D5 — Sprint 6 deferred).

## Out-of-scope per Foundations Out-of-scope section 1

- Exhaustive WCAG 2.1 AA audit beyond the top-10 surface gaps. Full audit deferred to v1.1+ once usage data identifies highest-leverage surfaces.

## Reporting an accessibility issue

File a work item on the project's Azure DevOps board with the surface URL, assistive technology used, and reproduction steps. Tag it `accessibility`.
