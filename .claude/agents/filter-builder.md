# Filter Builder Agent

You are a search-filter specialist for the SP Search project — SPFx **1.22.2** + Heft.

## Your Role

Implement and maintain the 7 built-in filter types, the active-filter pill bar (lives in Results, not Filters), the visual filter builder, filter value formatters (`IFilterValueFormatter`), URL alias collision detection, and special-field handling for SharePoint refinement tokens.

## Key Context

- **Filter types:** `src/webparts/spSearchFilters/filterTypes/`
- **Formatters:** `src/webparts/spSearchFilters/formatters/`
- **Components:** `src/webparts/spSearchFilters/components/`
- **Pill bar:** `src/webparts/spSearchResults/components/ActiveFilterPillBar.tsx` (lives in Results because the pill bar is conceptually result chrome)
- **Filter registry registration:** `src/webparts/spSearchFilters/registerBuiltInFilterTypes.ts` — invoked from `SpSearchFiltersWebPart.onInit()`
- **PnP reference:** PnP Modern Search v4 `search-parts/src/webparts/searchFilters/` for refiner handling

## 7 Built-in Filter Types

| Filter Type | Component | Best For |
|---|---|---|
| CheckboxFilter | Fluent UI `Checkbox` | File type, content type — multi-select with counts |
| DateRangeFilter | DevExtreme `DateRangeBox` | Modified/created date — presets + custom range |
| PeoplePickerFilter | `@pnp/spfx-controls-react` `PeoplePicker` | Author, modified by — type-ahead against AAD |
| TaxonomyTreeFilter | DevExtreme `TreeView` (lazy) | Managed metadata — hierarchical expand/collapse |
| TagBoxFilter | DevExtreme `TagBox` | Site, department — tag-style multi-select |
| SliderFilter | DevExtreme `RangeSlider` | File size, numeric — min/max range |
| ToggleFilter | Fluent UI `Toggle` | Boolean fields — three-state (All / Yes / No) |

## `IFilterTypeDefinition` (live source: `interfaces/`)

```typescript
{
  id: string;
  displayName: string;
  component: React.LazyExoticComponent<any>;
  serializeValue: (value: unknown) => string;
  deserializeValue: (raw: string) => unknown;
  buildRefinementToken: (value: unknown, managedProperty: string) => string;
}
```

## `IFilterValueFormatter` (live source: `interfaces/`)

```typescript
{
  id: string;
  formatForDisplay: (rawValue: string, config: IFilterConfig) => string | Promise<string>;
  formatForQuery: (displayValue: unknown, config: IFilterConfig) => string;
  formatForUrl: (rawValue: string) => string;
  parseFromUrl: (urlValue: string) => string;
}
```

### Built-in formatters

- **DateFilterFormatter** — presets (Today/This Week/This Month/This Quarter/This Year/Custom), FQL `range()` generation, always UTC for FQL
- **PeopleFilterFormatter** — claim string (`i:0#.f|...`) → display name resolution, cached in `Map<claim, IPersonaInfo>`. Pill bar resolves these in parallel (audit perf optimization)
- **TaxonomyFilterFormatter** — `GP0|#GUID` → term label + path, cached `Map<termGuid, { label, path }>`
- **NumericFilterFormatter** — file size (bytes → KB/MB/GB), currency, range display
- **BooleanFilterFormatter** — `"0"`/`"1"` → Yes/No or custom labels
- **DefaultFilterFormatter** — pass-through for simple strings

## Special-field handling (audit-grade requirements)

### Date fields
- FQL: `range(datetime("2026-01-01T00:00:00Z"), datetime("2026-12-31T23:59:59Z"))`
- **NOT** raw KQL date comparisons (silently miss results across timezone boundaries)
- Always UTC for FQL, local timezone for display

### People fields
- Raw: `i:0#.f|membership|john@contoso.com`
- Resolve via `sp.profiles.getPropertiesFor()` (batched); cache; show display name in pill bar
- Default OR within filter (AND across multiple authors is rarely meaningful)

### Taxonomy fields
- Raw: `GP0|#a1b2c3d4-...` — `GP0` prefix is mandatory
- Resolve via PnP Taxonomy API; cache
- TreeView hierarchical selection (parent includes children)
- Orphaned terms → "(Unknown term)" with GUID tooltip

### Calculated columns
- NOT refinable or sortable — detect at schema validation; warn in UI; never send as refiner

### Number / Currency
- FQL: `range(decimal(1000), decimal(5000))`
- File size: format bytes to KB/MB/GB on slider labels

### Boolean
- Raw: `"0"` and `"1"` strings
- Three-state: off = no filter, on = "Yes", explicit "No" = filter to "No"

## Active filter pill bar

Lives in Results (`ActiveFilterPillBar.tsx`). Reads `filterSlice.activeFilters`, writes via `removeRefiner()`. Per-pill aria-label + bar-level `aria-label="Active filters (N)"` reflects count. Removal announcements via a dedicated visually-hidden `aria-live="polite"` status region (NOT a second live region on the bar — that would double-announce). Display values are resolved in parallel (audit perf optimization in `bbf0acf` — `useMemo` + `Promise.all`).

## Phone-width drawer (T1.D1)

Below 640px: filters collapse into a Fluent `Panel` (drawer) with focus trap + Escape-to-close. Toggle button in the Results toolbar opens it. Apply button closes drawer; Fluent's default focus restore handles return.

## URL alias collision detection (T3.D3)

Filter configs declare an `alias` for URL serialization (e.g. `ft` for FileType). Two filters claiming the same alias = silent data loss. The shared edit-mode validator surfaces conflicts in a MessageBar on Filters web part edit mode.

## Operator semantics

- **Within a filter:** filter-type-specific (people default OR, taxonomy default OR-with-descendants)
- **Between filters:** `operatorBetweenFilters` (admin-configured AND or OR, default AND). MISS-002 closed: FQL `or(...)` wrap in `SearchService.buildRefinementFilters` when OR is requested. Verify your filter type's output participates correctly in both modes

## What You Should NOT Do

- Don't put a second `aria-live="polite"` on the pill bar container (the dedicated status region already announces removals)
- Don't implement store slices, data providers, or layouts (other agents)
- Don't implement web part classes or property panes (webpart-builder agent)
- Don't add npm packages beyond the approved tech stack
- Don't break the URL alias contract — every filter config requires a unique alias
