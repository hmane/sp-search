# SP Search — Theming and motion (T1.D9)

> Owned by T1 / Modern UI Quality. Covers the motion-reduction and
> theme-token contracts every shipped web part inherits.

## Motion reduction (`prefers-reduced-motion`)

Every web part's main SCSS module ships a `@media
(prefers-reduced-motion: reduce)` block at the bottom that clamps
transitions, animations, and `scroll-behavior` to 0.01ms. This is
applied via the universal selector pattern, so user-OS motion
preferences are honoured without per-surface opt-in:

```scss
@media (prefers-reduced-motion: reduce) {
  *,
  *::before,
  *::after {
    transition-duration: 0.01ms !important;
    transition-delay: 0ms !important;
    animation-duration: 0.01ms !important;
    animation-delay: 0ms !important;
    animation-iteration-count: 1 !important;
    scroll-behavior: auto !important;
  }
}
```

For per-surface animations that need a more nuanced fade-only-not-slide
behaviour, the shared `src/styles/motion.scss` mixin
`@include motion-safe` wraps a block so the inner styles only apply
when the user has NOT requested reduced motion. Surface-specific
animations (Filters drawer slide, Detail panel slide, Pivot ink-bar
transitions) all flow through this mixin.

Verified via `grep -rn "prefers-reduced-motion" src/styles src/webparts`
(≥6 hits across the 5 web parts + shared `motion.scss`).

## Dark-mode inheritance

SP Search does not ship its own dark theme. Instead, every web part
inherits the SharePoint Modern host's active theme via Fluent UI v8
theme tokens.

### Theme tokens consumed

Each `.module.scss` file uses the SPFx CSS-token + CSS-variable double
form so the same selector renders correctly under (a) the SharePoint
modern theme runtime (which injects `[theme:tokenName, default:#hex]`)
and (b) the Fluent UI Theme Provider runtime (which sets the matching
`--tokenName` CSS variable on the document root).

Example from `SpSearchResults.module.scss:21`:

```scss
.detailPanelNoPreviewCard {
  background-color: "[theme:neutralLighterAlt, default: #faf9f8]";
  background-color: var(--neutralLighterAlt, #faf9f8);
  border: 1px solid "[theme:neutralLighter, default: #f3f2f1]";
  border: 1px solid var(--neutralLighter, #f3f2f1);
}
```

When the SharePoint admin switches the site theme (e.g. to a dark
palette), Fluent UI's neutral / theme / status tokens all rotate, and
every web part picks up the new values without code changes.

### Tokens used by SP Search

The following Fluent UI v8 tokens are consumed across the suite —
admins can adjust them by editing the site theme via `Set-PnPWebTheme`
or the SharePoint admin centre.

**Neutral / surface tokens**:
- `white` — modal / panel background
- `neutralLighterAlt` — empty-state card, detail-panel surfaces
- `neutralLighter` — borders, dividers
- `neutralQuaternaryAlt` — disabled control fills
- `neutralTertiary` — subtle borders, icon defaults
- `neutralSecondary` — secondary text (.detailPanelDateLabel,
  shortcut help "shortcuts don't fire" hint)
- `neutralPrimary` / `bodyText` — body text
- `bodySubtext` — muted secondary text (preview-card subtitle)

**Theme tokens**:
- `themeLighterAlt` — selected-row backgrounds (List + Compact
  layouts)
- `themePrimary` — accent (links, primary buttons)
- `themeDarkAlt` — filter pill text

**Status tokens**:
- `errorText` / `errorBackground` — MessageBar error variant
- (Yellow / green for warning / success ride the Fluent
  MessageBar component itself, not custom CSS.)

### Limitations (v1.0)

- No standalone dark-mode override or auto-detection. The host
  SharePoint Modern theme drives everything.
- DevExtreme components (DataGrid, FilterBuilder, TreeView) have
  their own theme system — they DO NOT pick up SharePoint theme
  rotations automatically. If you change the site theme to a dark
  palette, DevExtreme surfaces stay on their default light theme.
  Setting `DataGrid.styles.colorScheme` per render via a Fluent
  theme listener is a future enhancement (T1.D9 follow-up).

### Verifying theme inheritance

1. Open the site theme in SharePoint admin → Change the look →
   pick a darker preset (e.g. "Dark Yellow").
2. Refresh the search page.
3. Confirm:
   - Page chrome turns dark (SharePoint host behaviour).
   - SP Search panels (Filters drawer, Detail panel) inherit the
     dark neutral surfaces.
   - Filter pills, links, and primary buttons rotate to the new
     theme palette.
   - DataGrid stays light (known limitation above).

The "3 before/after screenshots" the audit acceptance signal calls
out are captured manually as part of the smoke-checklist; the docs
above are the inheritance contract those screenshots verify.
