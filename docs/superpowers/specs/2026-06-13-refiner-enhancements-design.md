# Refiner Enhancements — Design

**Date:** 2026-06-13
**Status:** Draft — pending user review
**Scope:** 8 distinct refiner issues — 3 bugs + 5 enhancements

## Background

Production use of the post-1.0 SP Search solution surfaced 8 distinct issues with the Refiners surface. They share enough plumbing (filter slice, refiner mapping in `_mapRefiners`, the Filters web part's `handleToggleRefiner`) that fixing them piecemeal would mean re-touching the same files several times. This spec bundles them into one design and one implementation pass.

The issues, grouped by character:

| # | Label | Type | Severity |
|---|---|---|---|
| E | People refiner crashes with `Cannot read properties of undefined (reading 'defaultClass')` | Bug | High |
| A | Multi-value refiners (TagBox, PeoplePicker, TaxonomyTree, Dropdown-multi) clobber selections when more than one value changes at once | Bug | High |
| G | Refiner selection state doesn't visually persist after a search re-fires (most visible on Taxonomy) | Bug | Medium |
| B | Refiner values surface SharePoint's `string;#` / `int;#` / `datetime;#` raw storage prefix because no data-type awareness exists | Enhancement | Medium |
| F | Text refiners can't split delimited values (comma / semicolon / newline) into separate buckets | Enhancement | Medium |
| C | Taxonomy filter UI (DevExtreme `TreeView` with `searchEnabled`) reads as "search box with options below" — not a real tree | Enhancement | Medium |
| D | Toggle filter has no admin-configurable default value | Enhancement | Low-Medium |
| H | DevExtreme `TagBox` shows grey pills (default `dx.light.css`) because our Fluent-blue overrides lose the CSS specificity fight | Polish | Low |

## Goals

1. Eliminate the broken People picker (third-party module-css import crash) by moving off `@pnp/spfx-controls-react`'s `PeoplePicker` to Fluent UI v8 `NormalPeoplePicker`.
2. Eliminate the multi-toggle clobber bug — multi-value filters emit the **full intended selection** rather than looping per-delta.
3. Make refiner values **data-type aware**: strip `string;#`-style prefixes by default (auto-detect) with an admin override per refiner, and let admins split delimited Text values into separate buckets.
4. Replace the DevExtreme `TreeView` taxonomy filter with `@pnp/spfx-controls-react`'s `TaxonomyPicker`, with a webpack alias-shim fallback if its SCSS-module import crashes the same way `PropertyFieldCollectionData`'s did.
5. Add a configurable default value for the Toggle filter type.
6. Stop fighting DevExtreme's stock styling for `TagBox` — let `dx.light.css` win.

## Non-goals

- Defaults on filter types other than Toggle (Checkbox / DateRange / Slider / Text default-values are out of scope for this round).
- Replacing DevExtreme controls beyond removing the `TagBox` style overrides — the DataGrid / FilterBuilder / TagBox themselves stay as DevExtreme.
- Re-architecting the "stability merge" in `mergeRefiners`. It works as designed.
- Per-refiner formatter pipelines beyond what data-type metadata + value-split delivers. A full transform pipeline (regex, replace, format string) is deferred.

## Design

### 1. Fluent People picker (Issue E)

Replace `@pnp/spfx-controls-react`'s `PeoplePicker` with Fluent UI v8 `NormalPeoplePicker` inside [`PeoplePickerFilter.tsx`](../../../src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx).

- **Why Fluent and not the PnP fix-shim**: We already removed PnP's `PropertyFieldCollectionData` for the same broken module.scss reason. Each PnP control we keep is one more SCSS-shim we'd need to maintain. Fluent's `NormalPeoplePicker` is a first-party SPFx-safe primitive.
- **Claim resolution**: keep using `SPContext` + spfx-toolkit's existing user-claim resolution helpers. Wire `NormalPeoplePicker.onResolveSuggestions` to the same resolver the wrapper uses today.
- **Profile photos**: Fluent's `IPersonaProps.imageUrl` accepts the same `/_layouts/15/userphoto.aspx` URLs we already resolve.

**Out**: `@pnp/spfx-controls-react/lib/controls/peoplepicker/...` import from this filter (other surfaces that import `PeoplePicker` from PnP — if any — are out of scope here but should follow the same migration).

### 2. Batched selection callback (Issues A + G)

Today [`SpSearchFilters.tsx`](../../../src/webparts/spSearchFilters/components/SpSearchFilters.tsx) exposes `onToggleRefiner({ filterName, value, ... })` — single delta per call. Multi-value filters loop it, each call computes `buildNextFilters(stale_filters, delta)` against a stale React closure, and the second call's `setState` overwrites the first call's array. The last delta wins; the other deltas vanish.

**Fix**: Add a second callback `onReplaceRefinerValues({ filterName, values: IActiveFilter[] })` that accepts the **full intended selection** for a single filterName in one call. The parent computes:

```
nextActiveFilters = activeFilters.filter(f => f.filterName !== filterName).concat(values)
```

Migrate the 4 multi-value filters to this callback:
- `TagBoxFilter` — already has a `selectedTokens` array; emit it directly.
- `PeoplePickerFilter` — emit the resolved persona claims as one batch.
- `TaxonomyTreeFilter` — emit `selectedTokens` (the GP0|#GUID tokens) directly.
- `DropdownFilter` (multi-select mode) — emit `selectedKeys` as one batch.

Single-value filters (Checkbox, DateRange, Slider, Toggle, Text) keep `onToggleRefiner` unchanged.

**Why this fixes Issue G**: `selectedItemKeys` (TaxonomyTree) and equivalents derive from `activeFilters`. Once `activeFilters` actually reflects the user's full intent, the visual selection persists correctly on the next render.

### 3. Type-aware refiners + delimited splits (Issues B + F combined)

Add three new optional fields to [`IFilterConfig`](../../../src/libraries/spSearchStore/interfaces/IFilterTypes.ts):

```ts
interface IFilterConfig {
  // ... existing fields ...
  /**
   * Underlying SharePoint data type of the managed property.
   * 'auto' (default) runs the detect-and-strip heuristic in _mapRefiners.
   * Other values force the corresponding preprocessing.
   */
  dataType?: 'auto' | 'text' | 'choiceMulti' | 'lookup' | 'calculated' |
             'datetime' | 'yesno' | 'number';

  /**
   * If set, refiner values are split on this delimiter, trimmed,
   * deduplicated, and counts are aggregated per token. Useful for
   * Text columns that store comma/newline-separated tag-like values.
   */
  valueSplitDelimiter?: string;
}
```

Add a value-preprocessing pass to `_mapRefiners` in [`SharePointSearchProvider.ts`](../../../src/libraries/spSearchStore/providers/SharePointSearchProvider.ts). The pass needs access to the per-filter `IFilterConfig` — pass the active `filterConfig` array into `_mapRefiners` (currently it has none). For each `RefinementResults.Refiners.Entries[]`:

1. **Strip prefix** when `dataType` ∈ {choiceMulti, lookup, calculated} OR (`dataType === 'auto'` AND value matches `^[A-Za-z]+;#`):
   - `string;#Value` → `Value`
   - `int;#123` → `123`
   - `datetime;#2025-…` → `2025-…`
   - Empty-after-prefix entries (e.g. `string;#`) → emit as `(blank)` with original token preserved for KQL.
2. **Split** when `valueSplitDelimiter` is set: tokenize, trim each, drop empties, aggregate counts per token using a `Map<string, number>`. Preserve original raw value for the KQL clause (the SharePoint search will still match the raw multi-value field via `contains`-style semantics — the KQL clause uses the **token** not the prefix).
3. Return `IRefinerValue[]` where `name` is the cleaned display label and `value` is the KQL token.

**Auto-detect at runtime** (shipped): when `dataType` is left at the default `'auto'`, the preprocessing pass in `_mapRefiners` matches values against `^[A-Za-z]+;#` and strips the prefix when found. Admins can override per refiner via the property pane "Data format" section.

> **Deferred — schema-driven pre-fill**: an earlier draft proposed pre-filling `dataType` from `IManagedProperty.type` (YesNo / DateTime / Integer / Decimal) when admin picks a property via the schema helper. This was NOT shipped — the runtime heuristic + manual override cover the cases we've seen in practice. Adding schema-driven pre-fill is a candidate follow-up if admins routinely set `dataType` manually for the same property types.

**Property pane UI** ([`FiltersCollectionControl.tsx`](../../../src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx)):

Add a new "Data format" section visible when `filterType` ∈ {checkbox, tagbox, dropdown, text}:
- Dropdown: "Underlying data type" → `auto / text / choiceMulti / lookup / calculated / datetime / yesno / number`
- Text field: "Split values on" → admin types a delimiter (e.g. `,`, `;`, `\n`, `|`). Empty = no split.

Add the corresponding entries to [`fieldRelevance.ts`](../../../src/propertyPaneControls/filtersCollection/fieldRelevance.ts) so the section only appears for relevant filter types.

### 4. Taxonomy filter UI (Issue C) — **pivoted to DevExtreme TagBox**

> **History**: This section originally specified `@pnp/spfx-controls-react`'s `TaxonomyPicker`. We shipped it as commit `96e2264`, then pivoted after confirming the picker drops per-term refiner counts and cascade narrowing — both of which the project owner needs more than tree-browse. Final implementation in commit `7439e51` uses DevExtreme `TagBox` (mirrors `TagBoxFilter.tsx`) with PnP term-store label resolution. PnP `TaxonomyPicker` is no longer a project dependency for this surface.

Replace [`TaxonomyTreeFilter.tsx`](../../../src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx)'s DevExtreme `TreeView` with a DevExtreme `TagBox` implementation that mirrors [`TagBoxFilter.tsx`](../../../src/webparts/spSearchFilters/components/TagBoxFilter.tsx).

**Why TagBox not TaxonomyPicker**:
- Per-term refiner counts preserved — `values: IRefinerValue[]` already carries them, TagBox displays `"Label (count)"`.
- Cascade narrowing preserved — `values` IS the live refiner buckets, so other filters narrow the dropdown automatically.
- Selection persists synchronously from `activeFilters` (Issue G fixed).
- Tradeoff: no hierarchy display. Acceptable for flat taxonomies (the project owner's case); for genuinely hierarchical taxonomies, a follow-up could build a custom Fluent tree with counts overlaid.

**Label resolution**: SharePoint Search refiner buckets return GUIDs, not term labels. On mount, fetch labels via `SPContext.sp.termStore`:
- With `config.termSetId`: one call to `sets.getById(termSetId).terms()` returning the flat label dictionary.
- Without `config.termSetId`: per-GUID fan-out via `termStore.getTermById(guid)()`.

Cache in component state `Map<guid, label>`. Show a Fluent `Spinner` until first resolution completes; thereafter, cascade-introduced GUIDs resolve in the background without re-gating the UI.

**Selection wiring**: TagBox `onValueChanged` emits the next selected token array. Translate each `GP0|#<GUID>` token to an `IActiveFilter` via the new pure helper `buildTaxonomyTagBoxBatchPayload({ filterName, selectedTokens, labelMap, operator })`. Emit through `onReplaceRefinerValues` (Issue A pattern).

**Term set source**: keep the existing `config?.termSetId` field — same source.

### 5. Toggle default value (Issue D)

Add to `IFilterConfig`:

```ts
interface IFilterConfig {
  // ... existing ...
  /** Initial state of a 'toggle' filter when no URL / restored state is present. */
  defaultValue?: boolean;
}
```

Property pane:
- Add a `Toggle` control labelled "Default value" in the existing "Toggle labels" section of [`FiltersCollectionControl.tsx`](../../../src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx).
- Add `defaultValue` to the `'toggle'` case in [`fieldRelevance.ts`](../../../src/propertyPaneControls/filtersCollection/fieldRelevance.ts).

Seed logic in [`initializeSearchContext` at storeRegistry.ts:126](../../../src/libraries/spSearchStore/store/storeRegistry.ts#L126): after URL-restore runs, for each `filterConfig` where `filterType === 'toggle'` and `defaultValue !== undefined` AND no active filter exists for that `managedProperty`, push a synthetic `IActiveFilter` reflecting the default. URL-restore wins over defaults (preserves shareable links).

### 6. TagBox native DevExtreme styling (Issue H)

Remove the `.dx-tag` / `.dx-tagbox` / `.dx-list-item-selected` overrides at [`SpSearchFilters.module.scss:614-648`](../../../src/webparts/spSearchFilters/components/SpSearchFilters.module.scss#L614-L648). Keep the container layout (`.tagBoxFilterContainer`, `.sortControls`, `.showMoreBtn`).

DevExtreme's `dx.light.css` renders the stock grey pill — same as the DataGrid's tag cells already do today. Visual consistency within DevExtreme controls; accepted disconnect with Fluent-blue Checkbox/DateRange filters.

## Implementation order

Two PRs grouping related work:

**PR 1 — Bugs + small enhancements:**
1. Issue E (Fluent People picker)
2. Issue A + G (batched callback refactor across 4 filter components + parent)
3. Issue D (Toggle defaultValue)
4. Issue H (remove TagBox overrides)

Goal: ship the broken-thing fixes plus the small enhancement together. Reviewable in one sitting.

**PR 2 — Type-aware refiners + Taxonomy UI:**
5. Issue B + F (dataType + valueSplitDelimiter end-to-end)
6. Issue C (TaxonomyPicker swap, with pre-flight + alias-shim fallback)

Goal: separate PR because of property-pane schema additions + new KQL token transforms. Both deserve focused review.

## Testing

- **Unit**: Each filter component's "batched selection" path gets a Jest test that simulates 3 rapid value changes and asserts the resulting `activeFilters` contains all 3 (regression for Issue A).
- **Unit**: `_mapRefiners` gets a parameterised test for `dataType` strip + `valueSplitDelimiter` split combinations, including the empty-after-strip edge case.
- **Unit**: `initializeSearchContext` Toggle-default seed test — URL state present vs absent.
- **Integration**: Existing filter integration tests get extended with one cascading-narrow assertion per filter type to catch regressions in the cascade chain.
- **Manual**: Each filter type clicked through in the workbench against a fixture term set + multi-value Choice column to confirm display + KQL clauses.

## Out of scope (deferred)

- Defaults for non-Toggle filter types — out for this round; could be a small follow-up if asked.
- Replacing DevExtreme `TagBox` with a Fluent-native multi-select chip control — deferred until DevExtreme as a whole is reconsidered.
- A general-purpose refiner value transform pipeline (regex, custom JS) — type-aware preprocessing + split covers the cases on the table. We can add this later if real cases surface.

## Risks

| Risk | Mitigation |
|---|---|
| PnP `TaxonomyPicker` SCSS import crashes like `CollectionDataViewer` did | Pre-flight check + alias-shim fallback already in the design |
| `_mapRefiners` gaining a `filterConfig` dependency tangles a previously-pure mapper | Pass `filterConfig` as a function arg, not a service injection — keeps the mapper pure |
| Auto-detect strips legitimate values that happen to look like `type;#…` | Heuristic only fires when entry matches `^[A-Za-z]+;#` AND `dataType === 'auto'`; admins can override per-refiner |
| Toggle default values ride over user-cleared state on revisit | URL-restore wins over defaults — once user interacts, the URL has their selection and it sticks |
| Batched callback refactor introduces a subtle state-update bug | Migrate one filter at a time + run the cascade integration test between each |

## Files touched (estimate)

- `src/webparts/spSearchFilters/components/PeoplePickerFilter.tsx` (rewrite — Issue E)
- `src/webparts/spSearchFilters/components/TagBoxFilter.tsx` (batched callback — Issue A)
- `src/webparts/spSearchFilters/components/TaxonomyTreeFilter.tsx` (replaced — Issue C)
- `src/webparts/spSearchFilters/components/DropdownFilter.tsx` (batched callback — Issue A)
- `src/webparts/spSearchFilters/components/SpSearchFilters.tsx` (new `onReplaceRefinerValues` — Issue A)
- `src/webparts/spSearchFilters/components/SpSearchFilters.module.scss` (remove TagBox overrides — Issue H)
- `src/webparts/spSearchFilters/components/ToggleFilter.tsx` (read defaultValue from config — Issue D)
- `src/libraries/spSearchStore/interfaces/IFilterTypes.ts` (new fields — Issues B, D, F)
- `src/libraries/spSearchStore/providers/SharePointSearchProvider.ts` (`_mapRefiners` preprocessing — Issues B, F)
- `src/libraries/spSearchStore/store/storeRegistry.ts` (or equivalent init — Toggle defaultValue seed)
- `src/propertyPaneControls/filtersCollection/FiltersCollectionControl.tsx` (UI for new fields)
- `src/propertyPaneControls/filtersCollection/fieldRelevance.ts` (relevance map updates)
- `src/propertyPaneControls/pnpStyleShims/TaxonomyPicker.module.scss.js` (new file, conditional on pre-flight — Issue C)
- `gulpfile.js` (webpack alias, conditional — Issue C)
- Tests in `tests/webparts/spSearchFilters/` + `tests/providers/`
