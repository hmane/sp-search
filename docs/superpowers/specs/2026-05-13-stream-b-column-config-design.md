# Design — Stream B: layout-column configuration overhaul

> Status: approved 2026-05-13. Implementation lands in three phased commits — Phase 1 first; Phases 2 and 3 follow.
> Scope: covers items #2 (move `alias` off `selectedPropertiesCollection` onto the column collections), #3 (rich DataGrid + Compact column config with explicit renderers, multi-value separators, rich-text, max-length, see-more), and #4 (configurable column switcher) from the 11-item enhancement list.

## Problem

The DataGrid and Compact layouts' column-config is thin and the property-pane editor is already cramped:

- `gridPropertiesCollection` and `compactPropertiesCollection` items are `{ uniqueId, property }` — nothing else. No `alias`, no `width`, no visibility control, no choice of cell renderer (today's renderer is auto-detected by case-insensitive substring match against the property name — `title` / `author` / `date` / `fileType` / `fileSize` / `url` / fallback `text`).
- `alias` (the display label) lives only on `ISelectedPropertyItem` ([SpSearchResultsWebPart.ts:98-102](src/webparts/spSearchResults/SpSearchResultsWebPart.ts#L98)). That conflates two concerns: *which managed property to fetch* and *what to call the column in the UI*.
- All five collection editors (`selected`, `sortable`, `compact`, `grid`, `refinement`) use PnP's [`PropertyFieldCollectionData`](https://pnp.github.io/sp-dev-fx-property-controls/controls/PropertyFieldCollectionData/) — a flat tabular editor in a panel. Adding the 6-8 new fields per column item ([renderer, multi-value separator, rich-text, max-length, see-more, visibility, width, alias) would make that table unusable.

## Decision

### 1. New column-item schema (`IColumnConfigItem`)

Replaces `ILayoutPropertyItem` for `gridPropertiesCollection` and `compactPropertiesCollection`. `ISelectedPropertyItem` is unchanged in shape; its `alias` field becomes display-irrelevant (kept for back-compat — the new column-config item's `alias` wins).

```ts
type ColumnVisibility = 'always' | 'defaultOn' | 'defaultOff';

type ColumnRenderer =
  | ''            // empty = fall back to today's auto-detect (preserves current behaviour for migrated items)
  | 'text'
  | 'richText'
  | 'date'
  | 'number'
  | 'fileSize'
  | 'persona'
  | 'tags'
  | 'boolean'
  | 'url'
  | 'fileType';

type MultiValueSeparator = 'comma' | 'newline' | 'semicolon' | 'pill';

interface IColumnConfigItem {
  uniqueId: string;
  property: string;                       // managed-property name (must be in selectedProperties; provider already requests it via the baseline + admin's `selectedPropertiesCollection`)
  alias: string;                          // display label  ← #2 lives here now, off `selectedPropertiesCollection`
  width?: number;                         // px column width; 0 / undefined = auto
  visibility: ColumnVisibility;           // ← #4 — collapses show-by-default × in-chooser × always-on into one choice
  renderer: ColumnRenderer;
  maxLength?: number;                     // 0 / undefined = no truncation. Applies to text / richText / url renderers.
  seeMoreLink?: boolean;                  // when text is truncated, append a 'See more' affordance. Implies opening the detail panel / a tooltip on click.
  multiValueSeparator?: MultiValueSeparator;   // tags renderer only
}
```

`visibility` collapses three orthogonal concerns into one decision (admins don't need to grok two booleans):

| Value | Visible by default? | In the column chooser? |
|---|---|---|
| `always` | yes | no — protected (e.g. Title) |
| `defaultOn` | yes | yes — admin can toggle off |
| `defaultOff` | no | yes — admin can opt in |

### 2. Renderer set (built-in)

`renderer === ''` (empty) falls back to today's `resolveColumnKind` auto-detect — every migrated item starts in this state, so existing pages render byte-for-byte identically.

| `renderer` | Behaviour | Sub-fields revealed in editor |
|---|---|---|
| `''` *(empty)* | Auto-detect by property name (today's behaviour). | — |
| `text` | Plain text, escape-safe. | `maxLength`, `seeMoreLink` |
| `richText` | Sanitized HTML via spfx-toolkit's `htmlUtils/sanitizeHtml` (same pattern as `ListLayout` after Found.D4). | `maxLength`, `seeMoreLink` |
| `date` | Relative ("3 days ago") + absolute date in a tooltip — same `formatRelativeDate` + `formatDateTime` used elsewhere. | — |
| `number` | Locale-formatted number (`Intl.NumberFormat`). | — |
| `fileSize` | "1.2 MB" formatting on a raw byte count via `formatFileSize`. | — |
| `persona` | spfx-toolkit `UserPersona` (avatar + name + tooltip + Teams/email links). For columns with a single user value (e.g. `Editor`, `Author`, `CheckoutUser`). | — |
| `tags` | Multi-value (comma- or newline-split, taxonomy GP0 / pipe-separated user claims tolerated). Rendered as a list or pills per `multiValueSeparator`. | `multiValueSeparator` |
| `boolean` | ✓ / ✗ icon, no text — for fields like `IsDocument`, `EnableModeration`. | — |
| `url` | Clickable link, truncated to `maxLength` if set. | `maxLength` |
| `fileType` | File-type icon (existing `getFileTypeIconProps`) + uppercase extension label. | — |

> A registry pattern (`CellRendererRegistry` parallel to `LayoutRegistry` / `FilterTypeRegistry`) is **out of scope for v1**. The built-in set covers every shape we've seen in the Sprint-3 datagrids. If extension demand appears later, lift the dispatch into a registry — the call sites are already one function (`renderCell(item, column)`); no schema break needed.

### 3. Side-panel editor (the UX)

Replaces `PropertyFieldCollectionData` for `gridPropertiesCollection` and `compactPropertiesCollection` only. The other three collections (`selected`, `sortable`, `refinement`) keep PnP's collection-data — they are flat tables and don't need the panel.

**Compact list in the property pane** (per column):

- ↕ Drag handle / up–down arrows for reorder.
- A small "visibility" chip: 🔒 *Always* / 👁️ *Default-on* / 👁️‍🗨️ *Default-off*.
- `property` (the managed-property name, monospace) — read-only.
- `alias` (the display label) — read-only here, edited in the panel.
- A renderer chip ("Persona", "Date", "Tags", or "Auto" for empty).
- ✏️ `Edit` opens the side panel for that column.
- 🗑 `Remove` deletes it.
- "+ Add column" at the bottom — opens the panel in create-mode with a property picker.

**Side panel** (`@fluentui/react/lib/Panel`, `PanelType.medium`):

- **Header** — "Edit column: {alias || property}" or "Add column".
- **Property** — managed-property picker (the same `PropertyPaneSchemaHelper` we use for `queryTemplate`; falls back to free-text if the schema endpoint isn't available per its existing handling).
- **Alias** — text field (default = `property`).
- **Width** — number input (0 = auto).
- **Visibility** — choice group with three options.
- **Renderer** — dropdown of the built-in set + `Auto` for empty.
- **Renderer-specific options** — *revealed only when relevant*. `text` / `richText` / `url` → `maxLength` (number) + `seeMoreLink` (toggle). `tags` → `multiValueSeparator` (choice group: Comma / Newline / Semicolon / Pill).
- **Footer** — `Save` / `Cancel`.

The panel is rendered by a new component `ColumnConfigField` registered as a *custom* SPFx property-pane field (implements `IPropertyPaneField<IColumnConfigItem[]>`). The field's `render` mounts the compact list + the lazily-opened panel into the property-pane host.

### 4. Migration (zero-config upgrade)

Existing pages have `gridPropertiesCollection` / `compactPropertiesCollection` items shaped `{ uniqueId, property }`. They become `IColumnConfigItem` via `normalizeColumnConfigItem(raw)`:

```ts
function normalizeColumnConfigItem(raw: ILayoutPropertyItem | IColumnConfigItem): IColumnConfigItem {
  return {
    uniqueId: raw.uniqueId || generateUniqueId(),
    property: raw.property,
    alias: 'alias' in raw && raw.alias ? raw.alias : raw.property,
    width: 'width' in raw ? raw.width : undefined,
    visibility: 'visibility' in raw ? raw.visibility : 'defaultOn',
    renderer: 'renderer' in raw ? raw.renderer : '',   // '' triggers today's auto-detect
    maxLength: 'maxLength' in raw ? raw.maxLength : undefined,
    seeMoreLink: 'seeMoreLink' in raw ? raw.seeMoreLink : undefined,
    multiValueSeparator: 'multiValueSeparator' in raw ? raw.multiValueSeparator : undefined,
  };
}
```

`renderer: ''` is the migration sentinel — it routes through today's `resolveColumnKind` path, so existing pages render byte-for-byte identically until an admin opens the new editor and picks an explicit renderer.

`alias` migration (#2): `ISelectedPropertyItem.alias` stops affecting display. We don't remove the field (would break stored configs); we just stop reading it from the column-render path. `Manage Selected Properties` becomes purely "which managed properties does the search request?" — its `alias` column is preserved but documented as legacy (a one-line "this alias has moved to the Grid / Compact column editors" note in the property pane group description).

### 5. Apply scope

Both `gridPropertiesCollection` (DataGrid layout) and `compactPropertiesCollection` (Compact layout) adopt `IColumnConfigItem` and the new side-panel editor. The Compact layout already has a parallel `CompactColumnKind` auto-detect ([CompactLayout.tsx:23](src/webparts/spSearchResults/components/CompactLayout.tsx#L23)) — it gets re-pointed at the unified renderer dispatcher.

`titleDisplayMode` (the existing per-row Title rendering option) is unaffected — Title is always the first column and is rendered specially (anchor + DocumentTitleHoverCard + the Stream C / #7 `resolveResultLink` flow). It is one of the `visibility: 'always'` columns.

### 6. Column switcher (#4)

DataGrid's column chooser is gated by a new web-part property `showColumnChooser: boolean` *(default `true` — preserves today's affordance)*. When `true`, the chooser shows columns where `visibility !== 'always'`. Each column appears pre-checked iff its `visibility === 'defaultOn'`. Admins who want a fully fixed column set set `showColumnChooser: false` — the chooser disappears entirely.

## Phased implementation

The spec covers the whole design; the plan ships in three commits so each step is reviewable:

### Phase 1 — schema + migration + side-panel editor for `gridPropertiesCollection`

Shippable on its own; **does not change rendered DataGrid behaviour** (renderer set initially constrained to `''`-auto-detect + today's `ColumnKind` values exposed explicitly).

- New `src/webparts/spSearchResults/components/ColumnConfigField/` — the custom property-pane field (`ColumnConfigField.ts` field class + `ColumnConfigList.tsx` compact list + `ColumnConfigPanel.tsx` side panel). Unit-testable bits extracted to `columnConfig.ts` (the `normalizeColumnConfigItem` + the renderer-options-by-renderer table).
- New `IColumnConfigItem` interface (in `ISpSearchResultsProps.ts` or a new `columnConfig.ts`). Migration normalizer.
- `SpSearchResultsWebPart.ts` — `gridPropertiesCollection` switches from `PropertyFieldCollectionData` to the new field; `_getGridPropertyColumns()` returns the migrated items; `IWebPartProps.gridPropertiesCollection: IColumnConfigItem[]`.
- `DataGridContent.tsx` — `resolveColumn(item, column)` consults `column.renderer` first, falls back to today's auto-detect when empty. New props from `IColumnConfigItem` thread through.
- Tests for `normalizeColumnConfigItem` (every field default + every renderer accepted).

### Phase 2 — new renderer types

- `renderCell.ts` — central renderer dispatcher with one function per renderer type. Replaces the inline `cellRender` switch in `DataGridContent.tsx`.
- New renderers: `richText` (sanitized HTML), `persona` (UserPersona), `tags` (multi-value with separator), `boolean` (icon), plus a unified `number` / `fileSize` (Intl-formatted) — the explicit equivalents of today's auto-detected kinds.
- Side-panel editor reveals the renderer-specific sub-fields (`maxLength` / `seeMoreLink` / `multiValueSeparator`) when the relevant renderer is chosen.
- Tests for each renderer's output (snapshot-light: render to string, assert key substrings; the full DOM tests are integration-level and out of scope).

### Phase 3 — Compact layout + `#4` column-chooser config

- `compactPropertiesCollection` switches to the new field + the renderer dispatch; `CompactColumnKind` is retired (becomes a thin wrapper around the unified dispatcher).
- New web-part property `showColumnChooser: boolean` (default `true`); DataGrid wires it to its column-chooser feature; manifest defaults updated.
- Column chooser respects `visibility === 'always'` (those columns never appear in the chooser) and pre-checked state matches `visibility === 'defaultOn'`.
- `ISelectedPropertyItem.alias` quiet-deprecation: stop reading it from the column-render path; add a single-line description in the property-pane group ("Display labels moved to the Grid / Compact column editors").

## Verification

Each phase:

- `npm run type-check` clean.
- `npm test` green; new tests added per phase (column-config normalizer in Phase 1; per-renderer output in Phase 2; column-chooser visibility behaviour in Phase 3).
- `npm run package` clean (no new Sass deprecation warnings).
- `npm run check:bundles` green — Phase 1's new property-pane component adds ~3–5 KB to the Results bundle; well within the 1.1 MB budget. Phase 2's renderer dispatcher adds ~2 KB. Phase 3 is mostly wiring (~0.5 KB).
- Manual: open the property pane, edit a column, verify the side-panel UX renders correctly; reorder columns; toggle visibility; verify migration of an existing page leaves rendering identical.

## Out of scope (deferred)

- A `CellRendererRegistry` extension point (parallel to `LayoutRegistry` / `FilterTypeRegistry`). Built-in set only for v1; add the registry if/when a custom-renderer use case appears.
- Per-column conditional visibility (visible only when a particular filter is active, etc.). Stream D / #5 tackles audience-targeted *refiners*; the parallel for *columns* is a later request.
- The `selectedPropertiesCollection` editor itself — kept as-is; only the role of its `alias` field changes (quietly).
- Image preview in `ColumnConfigPanel` — admins editing an image-column see only a renderer label, not a sample image.
- Drag-and-drop preview / ghost while reordering. Reorder uses up/down arrows + keyboard accessible; HTML5 DnD is a follow-up.
