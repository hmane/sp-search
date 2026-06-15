# DataGrid Token / Badge Cell Rendering — Design

**Date:** 2026-06-15
**Status:** Approved — pending implementation plan
**Scope:** One enhancement to the DataGrid layout's per-column rendering — configurable value splitting plus token/badge display with admin value→color mapping and auto-color fallback.

## Background

The DataGrid layout renders each column through a per-column renderer chosen by an admin in the column-config side panel (`ColumnConfigPanel.tsx`). The renderer set lives in [`renderCell.tsx`](../../../src/webparts/spSearchResults/components/renderCell.tsx) and the per-column config shape in [`columnConfig.ts`](../../../src/webparts/spSearchResults/components/ColumnConfigField/columnConfig.ts). Today there are 11 renderers: auto-detect, `text`, `richText`, `number`, `fileSize`, `date`, `boolean`, `tags`, `persona`, `url`, `fileType`.

The `tags` renderer already splits a value and can show it inline, one-per-line, or as neutral grey "pill" chips (`multiValueSeparator: comma | semicolon | newline | pill`). Two limitations make it fall short of common needs:

1. **The split character is hardcoded** to `,` / `;` (`splitTagValue` in `renderCell.tsx`). A field delimited by `|`, a newline, or any other character can't be broken into separate values.
2. **Pills are always a flat neutral grey.** Categorical and status-style values (Approved / Pending / Rejected, Low / Medium / High, department names) carry meaning that a colored badge would convey at a glance, but there's no way to color a token.

This spec closes both gaps with the smallest change that fits the existing per-column model: a **configurable split delimiter** and a **colored badge display style** whose colors come from an optional admin value→color map, falling back to a deterministic auto-color for any unmapped value.

## Goals

1. Let an admin set the **character(s) a column's value is split on**, decoupled from how the split values are displayed.
2. Add a **"Badges (colored)" display style** to the `tags` renderer that renders each split value as a styled chip.
3. Let an admin **map specific values to named colors** (and optionally an icon), e.g. `Approved → green`, `Overdue → red`.
4. **Auto-color any unmapped value** deterministically (same value always gets the same color) so a column looks good with zero or partial mapping.
5. Keep everything **theme-safe and accessible** (light/dark variants, WCAG-AA text contrast) and **back-compatible** (existing `tags`/`pill` columns render unchanged).

## Non-goals

- Numeric visualizations (progress bars, star ratings, data bars, trend arrows) — a separate future spec.
- A conditional-formatting rules engine (value/condition → text color/weight/row tint across arbitrary renderers) — a separate future spec.
- `{property}` composite templates that compose a cell from multiple fields — explicitly dropped (the word "token" here means a badge chip, not a placeholder).
- Free-form hex color entry. Colors are chosen from a curated named palette so theme variants and contrast stay correct.
- Changing any renderer other than `tags`, or the auto-detect path.

## Design

### 1. Renderer surface — enhance `tags`, relabel to "Tags / badges"

No new renderer type. The existing `tags` renderer (internal key stays `'tags'` for back-compat) gains the new behavior. Its label in the editor dropdown changes to **"Tags / badges"** so it's discoverable for single-value status columns — a single value simply splits into one token, which renders as one badge.

`DataGridContent.tsx` already dispatches `tags` → `renderTags` and passes the full `IColumnConfigItem` to the renderer, so no dispatch change is needed beyond letting the new fields flow through. `kindFromRenderer` keeps `tags → { kind: 'tags', minWidth: 160 }`.

### 2. Config model — new optional fields on `IColumnConfigItem`

Added to [`columnConfig.ts`](../../../src/webparts/spSearchResults/components/ColumnConfigField/columnConfig.ts):

```ts
export type BadgeColor =
  | 'neutral' | 'blue' | 'teal' | 'green'
  | 'amber'   | 'orange' | 'red' | 'purple' | 'magenta';

export interface IBadgeColorRule {
  value: string;        // matched case-insensitively against the cleaned cell value
  color: BadgeColor;
  icon?: string;        // optional Fluent icon name, rendered before the label
}

// IColumnConfigItem additions (all optional → back-compat):
splitDelimiter?: string;            // literal delimiter; "\n"/"\t" tokens → whitespace; unset → today's , / ;
multiValueSeparator?: ... | 'badge';  // add 'badge' to the existing union
valueColorMap?: IBadgeColorRule[];  // admin value→color rules; only used when display = 'badge'
autoColorUnmapped?: boolean;        // default true; only used when display = 'badge'
```

- `AUTO_COLOR_PALETTE` = the 8 non-neutral `BadgeColor`s, used for hashing.
- `multiValueSeparator` continues to mean "display style": `comma` / `semicolon` = inline join, `newline` = stacked, `pill` = neutral chips, **`badge` = colored chips (new)**.
- `valueColorMap` / `autoColorUnmapped` are ignored unless `multiValueSeparator === 'badge'`.

**Normalizer** (`normalizeColumnConfigItem`) additions:
- `splitDelimiter`: keep if a non-empty string ≤ 8 chars, else `undefined`.
- `multiValueSeparator`: add `'badge'` to `VALID_SEPARATORS`.
- `valueColorMap`: keep array entries that have a non-empty `value` and a `color` in the `BadgeColor` set; trim values; **dedupe by lowercased value** (first wins); cap at 50 entries; drop `icon` if not a non-empty string.
- `autoColorUnmapped`: keep boolean; default `true` when `multiValueSeparator === 'badge'` and unset, else preserve as given / `undefined`.

### 3. Color resolution (pure helper in `renderCell.tsx`)

For a badge column, build a lowercased lookup `Map<string, IBadgeColorRule>` from `valueColorMap` once per render, then resolve each token:

1. **Mapped** — case-insensitive exact match on the cleaned, trimmed token → that rule's color (+ icon).
2. **Auto** — else if `autoColorUnmapped !== false` → `AUTO_COLOR_PALETTE[hash(tokenLowerTrimmed) % palette.length]`. Deterministic, no state, stable across rows and re-renders.
3. **Neutral** — else → `'neutral'`.

Hash:
```ts
function hashBadgeColorIndex(s: string, len: number): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) { h = (h * 31 + s.charCodeAt(i)) | 0; }
  return Math.abs(h) % len;
}
```

### 4. Rendering (`renderCell.tsx`)

- **Configurable split.** Generalize `splitTagValue` to accept the column's `splitDelimiter`. If set: interpret literal `\n`→newline and `\t`→tab, regex-escape the rest, and split on `/\s*<delim>\s*/`. If unset: today's `/\s*[,;]\s*/`. SharePoint prefix cleaning (`cleanSearchResultDisplayText`) is still applied to the whole value and each part, exactly as today.
- **Badge display.** When `multiValueSeparator === 'badge'`, map each token to a chip: optional leading icon (`<Icon>`, `aria-hidden`), then the label text, with `title` = full token. The chip's CSS class is `styles.gridBadge` + `styles['gridBadge--' + color]`.
- Existing `comma` / `semicolon` / `newline` / `pill` paths are untouched.
- Empty value → `muted()` ("--"), as today.

### 5. Palette + theming (`SpSearchResults.module.scss`)

- One base `.gridBadge` class (inline-flex chip: small radius, padding, font-size, gap for icon) plus nine `--<color>` modifier classes.
- Each color modifier sets a **soft background + readable foreground** pair chosen for AA contrast on both light and dark section backgrounds (soft translucent fill so the chip reads against themed sections). `neutral` reuses the existing pill greys.
- Colors are fixed, accessible pairs (not `[theme:...]` accents) so the nine options stay visually distinct and legible regardless of the site's theme accent.

### 6. Edit-mode UX (`ColumnConfigPanel.tsx`)

When the renderer is `tags` (Tags / badges), show, in addition to today's fields:

- **Split character** — `TextField`, placeholder `;  ,  |`, help text noting `\n` = new line. Bound to `splitDelimiter`.
- **Display style** — the existing `SEPARATOR_OPTIONS` dropdown gains **"Badges (colored)"** (`badge`). (Existing: Comma-separated / One per line / Semicolon-separated / Pills.)
- **When display = Badges only:**
  - **Value colors** — a compact mapping editor: rows of *Value* (`TextField`) · *Color* (`Dropdown` of the 9 named colors, each option showing a color swatch) · *Icon* (optional `TextField` for a Fluent icon name), with add/remove-row controls. Bound to `valueColorMap`.
  - **Auto-color other values** — `Toggle`, default on. Bound to `autoColorUnmapped`.

The mapping editor is a small self-contained sub-component within the panel; it edits the `valueColorMap` array immutably through the existing `update(...)` path. New field labels follow `ColumnConfigPanel`'s current pattern — **hardcoded English literals** (the panel does not use loc strings today), so no loc file change is needed.

### 7. Accessibility

- Color is **decorative**; the visible text label carries the meaning, so screen readers announce the value regardless of color. Icons are `aria-hidden`.
- Every color modifier meets WCAG-AA (≥ 4.5:1) text-on-chip contrast.
- Full token text is available on hover via `title`.

### 8. Testing

`renderCell.test.tsx`:
- Custom `splitDelimiter` (`|`, `\n`) splits correctly; default `,`/`;` unchanged.
- SharePoint `string;#` prefix cleaning still applies to whole value and per-part.
- Color resolution: mapped exact + case-insensitive; auto-color **stable** (same value → same color across calls) and **distinct-ish** across different values; `autoColorUnmapped: false` → neutral; unmapped with auto on → a non-neutral palette color.
- Badge chips render label + optional icon; `title` carries the full value; empty → muted dash.

`columnConfig.test.ts`:
- Normalizer round-trips `splitDelimiter` (kept / capped / dropped), `multiValueSeparator: 'badge'`, `valueColorMap` (invalid color dropped, value trimmed, dedupe by lowercased value, cap), and `autoColorUnmapped` default.

## Files touched

| File | Change |
|---|---|
| `components/ColumnConfigField/columnConfig.ts` | `BadgeColor`, `IBadgeColorRule`, palette constant, new `IColumnConfigItem` fields, normalizer + validation |
| `components/renderCell.tsx` | configurable split, color-resolution helper + hash, badge chip rendering |
| `components/SpSearchResults.module.scss` | `.gridBadge` base + 9 color modifier classes |
| `components/ColumnConfigField/ColumnConfigPanel.tsx` | split field, "Badges" display option, value-color map editor, auto-color toggle (hardcoded labels, per existing panel pattern) |
| `tests/webparts/spSearchResults/renderCell.test.tsx` | badge + split tests |
| `tests/webparts/spSearchResults/columnConfig.test.ts` | normalizer tests for new fields |

`DataGridContent.tsx` needs no change beyond the new fields flowing through the already-passed column object. The Experience web part inherits the behavior automatically (it renders the same Results component and column config).

## Resolved decisions

- **Enhance `tags`** rather than add a separate `badge` renderer — avoids duplicating split logic and a new auto-detect entry; one renderer serves both multi-value tags and single-value status badges.
- **Named palette**, not raw hex — keeps theme variants and contrast correct.
- **Per-value icons included** in Phase 1 as an optional field — small addition, natural for status chips; an admin may leave it blank.

## Out of scope (future specs)

- Numeric visualizations (progress / rating / data-bar / trend).
- Conditional-formatting rules engine.
- `{property}` composite/templated cells.
