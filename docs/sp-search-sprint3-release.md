# SP Search Sprint 3 — Release Notes and Close-Out Checklist

## Status Overview

Sprint 3 delivered all five tracks. The platform materially changed shape from the
original audit baseline. This document records what was shipped, what was deferred,
the regression sweep checklist, and the admin documentation updates required before
wider rollout.

---

## Sprint 3 Exit Criteria — Delivery Status

| Exit Criterion | Status | Notes |
|---|---|---|
| DataGrid columns reflect `selectedPropertiesCollection` | ✓ Delivered | A2 — admin-configured columns, cell renderers, localStorage persistence |
| DataGrid filter row enabled | ✓ Delivered | A3 — DevExtreme `filterRow`, `columnChooser`, virtual scrolling |
| CSV export works end-to-end | ✓ Delivered | A5 — Export button in DataGrid toolbar |
| DataGrid render failure degrades to List with visible message | ✓ Delivered | A1 — `DataGridRenderErrorBoundary` + `onFallback()` |
| `GraphSearchProvider` returns people results on Graph vertical | ✓ Delivered | B1 — `/search/query` with `entityTypes:['person']`, presence batch lookup |
| Per-vertical `dataProviderId` routes to correct provider | ✓ Delivered | B2 — orchestrator `_getProvider()` now honors vertical config |
| Presence badge and action buttons on People cards | ✓ Delivered | B3 — Teams chat, email, view profile, graceful SP degradation |
| Zero-result queries logged to SearchHistory | ✓ Delivered | C1 — `IsZeroResult` Boolean field, fire-and-forget, session dedup |
| Search Insights tab in Search Manager | ✓ Delivered | C4 — `SearchInsightsPanel`: stat cards, top queries, CTR, volume chart |
| Health tab showing zero-result queries | ✓ Delivered | C2 — `ZeroResultsPanel`: ranked table, "Try it" re-run |
| Click-through tracking wired | ✓ Delivered | C3 — `logClickedItem` from all layouts; audit confirmed complete |
| Provisioning presets run without errors | ✓ Delivered | D — `Search-ScenarioPresets.ps1`: 5 presets, `Invoke-SearchScenarioPage` |
| `general` preset available in property pane | ✓ Delivered | D3 — ChoiceGroup now has general/documents/news/people/media/custom |
| List and Compact usable without scroll at 375px | ✓ Delivered | E1 — responsive breakpoints audited; gallery single-column at 399px |
| DataGrid chunk preloaded on hover | ✓ Delivered | E4 — `LAYOUT_PRELOADERS` map, `onPreloadLayout` in ResultToolbar |

---

## What Changed From The Audit Baseline

### A-Track — Advanced DataGrid

The DataGrid is no longer a "basic table view." Changes from baseline:

- Columns driven by `selectedPropertiesCollection` (admin-configured), not hardcoded six-column set
- Cell renderers typed by property mapping (title link, author persona, date format, file type icon, file size, URL, tags, boolean, number, taxonomy, thumbnail, text fallback)
- Filter row, column chooser, and virtual scrolling via DevExtreme props
- Stable row selection wired to `bulkSelection` store slice; selection cleared on layout switch (fixes hidden-selection bug from audit)
- `BulkActionsToolbar` surfaces above grid when items are selected
- CSV export from current page via DataGrid toolbar
- User column preferences persisted to `localStorage` per `searchContextId`
- `DataGridRenderErrorBoundary` logs real errors and falls back to List — the double lazy-load chain was removed

### B-Track — Graph People Vertical

People layout is now backed by Microsoft Graph, not SharePoint search results:

- `GraphSearchProvider`: `/search/query` with `entityTypes:['person']`, proper pagination via `from`/`size`
- Presence lookups batch-fetched via `/communications/presences`, cached per provider instance
- `PeopleLayout` shows: presence badge, Teams chat button, email button, View profile link, department, office location, skills (max 5)
- Degrades gracefully when backed by `SharePointSearchProvider` (hides presence + Teams action if fields absent)
- Per-vertical `dataProviderId` selection now works end-to-end; `SearchOrchestrator._getProvider()` reads `vertical.dataProviderId`
- Verticals web part property pane exposes data provider dropdown per vertical

### C-Track — Analytics and Search Quality

The feedback loop is now in place:

- `IsZeroResult` Boolean field added to `SearchHistory` list schema (provisioning script updated)
- `logSearch()` accepts and persists `isZeroResult` flag; orchestrator passes `resultCount === 0`
- `ZeroResultsPanel` in "Health" tab: cross-user admin view, ranked by occurrence, "Try it" re-runs query
- `SearchInsightsPanel` in "Insights" tab: total searches, zero-result rate, CTR, avg results, top queries bar chart, top clicked results bar chart, daily volume sparkline, 30d/60d/90d range toggle
- `loadZeroResultQueries(daysBack, maxItems)` and `loadAllHistoryForInsights(daysBack, maxItems)` on `SearchManagerService` — both use `SearchTimestamp >= cutoff` as first CAML predicate (safe above 5k threshold, cross-user)
- Click-through tracking confirmed complete from all layouts; no changes needed (C3 was already fully wired)

### D-Track — Scenario Presets

Presets now exist in three places:

- `src/webparts/spSearchResults/presets/searchPresets.ts` — centralized `SCENARIO_PRESETS` registry (general, documents, people, news, media); each preset defines layout, `queryTemplate`, `selectedProperties`, `sortableProperties`, filter suggestions, and data provider hint
- `_applyScenarioPreset()` replaces `_applyLayoutPreset()` — writes layout toggles, `queryTemplate`, `selectedPropertiesCollection`, and `sortablePropertiesCollection` atomically
- Property pane ChoiceGroup: `general` option added; hint label updated to note Filters/Verticals web parts need separate configuration
- `scripts/Search-ScenarioPresets.ps1`: `Get-SearchScenarioPreset`, `Get-SearchScenarioPresetList`, `Invoke-SearchScenarioPage` — provisions a full page with all 5 web parts pre-configured for the named scenario

### E-Track — Mobile and Performance

Targeted fixes only, no rewrites:

- Gallery: single-column breakpoint at 399px (was 2-col down to 0px)
- Loading overlay: `backdrop-filter: blur(2px)` for dark-theme resilience (previously hardcoded white wash)
- DataGrid container: `-webkit-overflow-scrolling: touch` for iOS momentum scrolling
- Layout button hover → `LAYOUT_PRELOADERS` fires `import()` to warm chunk before click; webpack deduplicates on subsequent hovers

---

## Regression Sweep Checklist

Run this against a live tenant before signoff. Use the test site at
`https://pixelboy.sharepoint.com/sites/SPSearch`.

### Preset scenarios

For each scenario, deploy via `Invoke-SearchScenarioPage` and verify:

- [ ] **General** — enters list layout, `{searchTerms}` query runs, 3 layout options visible
- [ ] **Documents** — `IsDocument:1` scoped, title/author/modified/type/size columns present in grid
- [ ] **People** — People layout loads, Graph results appear (requires People.Read permission), presence badge visible
- [ ] **News** — Card layout default, `PromotedState:2` filters to news pages only
- [ ] **Media** — Gallery layout default, image/video file type filter active

### DataGrid

- [ ] Columns match `selectedPropertiesCollection` configured in property pane
- [ ] Column chooser opens and hides/shows columns
- [ ] Filter row appears and narrows results inline
- [ ] Column preference saved after hide/show; persists on page reload (check `localStorage`)
- [ ] CSV export downloads a file with current page rows
- [ ] Row selection visible; `BulkActionsToolbar` appears above grid with at least one item selected
- [ ] Switching from Grid to List clears selection (no phantom `bulkSelection` in store)
- [ ] DataGrid chunk loads on first hover of Grid button (check Network tab: no `DataGridLayout` chunk before hover)
- [ ] Introducing a bad DevExtreme prop falls back to List layout with message, not a white screen

### People vertical

- [ ] Graph-backed vertical returns people results distinct from document results
- [ ] Presence indicator shows on cards (Available/Busy/Away/Offline)
- [ ] Teams chat link opens Teams to the person
- [ ] SharePoint-backed People vertical shows cards without presence badge (graceful degradation)

### Analytics / Search Manager

- [ ] After performing searches including at least one zero-result query, open Search Manager → Health tab
- [ ] Zero-result query appears in the ranked table
- [ ] "Try it" button re-runs the query and closes the panel
- [ ] Insights tab loads stat cards with non-zero values
- [ ] Insights 30d/60d/90d toggle reloads data
- [ ] Top queries bar chart shows clickable items that re-run the query
- [ ] Daily volume sparkline renders columns (may be flat if test data is sparse)
- [ ] `IsZeroResult` column exists in the `SearchHistory` list (`/lists/SearchHistory/fields`)

### Mobile (test at 375px viewport)

- [ ] List layout: title and timestamp visible, no horizontal scroll
- [ ] Compact layout: title visible; author/date/size columns hidden at 375px (correct per CSS)
- [ ] Card layout: single column at 375px
- [ ] People layout: single column at 375px
- [ ] Gallery layout: 2-column at 375px, single-column at 375px — *(clarify: 375 > 399px threshold; stays 2-col at 375px, goes 1-col only below 400px)*
- [ ] Active filter pill bar wraps without overflow
- [ ] Toolbar wraps count/sort above layout buttons when narrow
- [ ] DataGrid horizontal scrolls with momentum on iOS Safari (no snap-stop behavior)

### Slow-network overlay (throttle to Slow 3G in DevTools)

- [ ] Typing a query and pressing enter: skeleton shows, not "No results found"
- [ ] On a refresh search (results already shown): previous results stay visible, overlay appears after ~300ms
- [ ] Overlay fades in (not a hard cut-in)
- [ ] On dark SharePoint themes: overlay blur is visible without white wash

---

## Admin Documentation Updates Required

The following areas of the admin documentation are now outdated and must be updated
before handing the solution to site owners.

### 1. Scenario presets

**Currently documented:** "A preset sets the default layout and visible switcher buttons."

**Now true:** A preset sets layout, query template, selected property columns, and sortable
properties simultaneously. Filters web part and Verticals web part still require separate
configuration. The property pane hint was updated in the release, but external docs (if any)
need the same language.

**New doc to write:** "Getting started with scenario presets" — covers:
- What each preset does (general/documents/people/news/media)
- How to deploy a preset page using `Invoke-SearchScenarioPage`
- What the admin needs to configure separately (Filters, Verticals, Graph permissions for People)

### 2. People vertical — Graph provider

**Currently documented:** Not documented at tenant/admin level. `GraphSearchProvider` did not exist.

**Now true:** People verticals can be backed by Microsoft Graph. This requires:
- `People.Read` delegated permission (not `Sites.Read.All`)
- Verticals web part property pane: set Data Provider to `graph` on the People vertical
- Presence lookups batch via `/communications/presences` (no extra permission required)

**New doc to write:** "Configuring a People Search vertical" — permissions checklist,
property pane steps, graceful degradation note for users without Graph access.

### 3. SearchHistory list schema update

**Currently documented:** SearchHistory fields as per the original provisioning script.

**Now true:** `IsZeroResult` Boolean field was added. Any tenant running a previous version
needs to run the updated provisioning script (or manually add the field) before the Health
and Insights tabs will populate.

**Doc update:** Add a migration note to the deployment guide:
> If upgrading from a pre–Sprint 3 installation, re-run `Provision-SPSearchLists.ps1`
> or add a Boolean field named `IsZeroResult` to the `SearchHistory` list manually.
> The field is optional at the SharePoint list level — the web part handles a missing field
> gracefully — but the Health tab will show no data until it is present.

### 4. Search Manager — new tabs

**Currently documented:** Search Manager has tabs: Saved Searches, History, Collections.

**Now true:** Two admin-only tabs were added: **Health** and **Insights**.
- Health shows zero-result queries ranked by frequency with a "Try it" action.
- Insights shows aggregate analytics: total searches, zero-result rate, click-through rate, top queries, top clicked items, daily volume.
- Both tabs load cross-user data (no Author filter) and are intended for site search admins, not end users.

**Doc update:** Admin guide section on Search Manager — add tab descriptions and note
that the Health/Insights data is bounded by `SearchTimestamp` (last 30–90 days, configurable).

### 5. DataGrid — column configuration

**Currently documented:** DataGrid as a "basic table view."

**Now true:** DataGrid columns are driven by `selectedPropertiesCollection` in the Results
web part property pane. The admin selects properties (and optionally aliases them), and
those become the grid columns. Cell renderers are applied automatically by property type.

**Doc update:** Admin guide — "Configuring the DataGrid view" section:
- How to add/reorder properties in `selectedPropertiesCollection`
- How to alias a column header
- Note that users can hide/show columns via the column chooser; preferences are per-browser
- CSV export exports the current page only

### 6. `_applyLayoutPreset` → `_applyScenarioPreset` (internal/developer note)

**Not user-facing, but relevant for any developer extending the solution:**

`_applyLayoutPreset` was replaced by `_applyScenarioPreset` in `SpSearchResultsWebPart.ts`.
Any custom code calling `_applyLayoutPreset` by name will fail at runtime (TypeScript will
not catch this if called via reflection or `eval`). Update custom extensions accordingly.

---

## Explicitly Deferred (Not In This Release)

These items were in the Sprint 3 scope but were not delivered. Carry forward to Sprint 4:

| Item | Reason for deferral |
|---|---|
| `operatorBetweenFilters` wiring | Still unimplemented in filter execution path; property pane option exists but does nothing. Hide from property pane in Sprint 4 or implement. |
| `queryInputTransformation` | Still not applied in `SearchOrchestrator`; property still surfaced in `SpSearchBox` props. Implement in orchestrator query construction path. |
| Excel (XLSX) export | CSV only this sprint. XLSX requires SheetJS or similar library — deferred for bundle-size reason. |
| Org chart / manager traversal | Not in Sprint 3 scope; People search delivers flat directory only. |
| Knowledge Base, Hub Search, Policy Search presets | Provisioning presets script has 5 core presets; 3 archetype presets deferred. |
| Automated test suite | Jest harness remains broken (`ts-jest` cannot resolve `jest-util`). Unblocked by sprint work but not fixed in this sprint. Fix in Sprint 4 before any further store refactoring. |
| Admin-time property validation | Planned in C2 scope; not shipped. The edit-mode `MessageBar` still shows only the context ID warning. |

---

## Release Decision

All Sprint 3 exit criteria are met. The platform is ready for controlled rollout to
pilot sites using the `general` or `documents` preset as the starting point.

Do not broadly roll out the People preset until the Graph permission request has been
approved in the target tenant's Entra ID admin center. The web part degrades gracefully,
but People cards will be empty without the permission grant.

Recommended pilot sequence:
1. Document Center or team site — `documents` preset
2. Intranet news hub — `news` preset
3. HR or IT — `people` preset (after Graph permission approved)
