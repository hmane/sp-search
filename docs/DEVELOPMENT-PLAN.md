# SP Search — Development Plan & Progress Tracker

**Version:** 1.0
**Created:** February 5, 2026
**Requirements Reference:** [sp-search-requirements.md](./sp-search-requirements.md) v1.4

---

## How to Use This Document

- Each task has a checkbox `[ ]` — mark `[x]` when complete
- Tasks within a step are ordered by dependency — complete top-to-bottom
- **Agent** column references which `.claude/agents/` specialist handles the task
- **Req §** points to the requirements section for full spec details
- Steps marked **GATE** must be fully complete before the next step begins

---

## Phase 1: Foundation

> **Goal:** Scaffolded SPFx solution with working end-to-end search — type a query, get results, refine with basic filters, switch verticals. No fancy layouts, no saved searches, no detail panel.

---

### Step 1.0 — Project Scaffolding (GATE)

> SPFx project structure, dependencies, build pipeline.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.0.1 | Scaffold SPFx 1.21.1 solution with Yeoman (`yo @microsoft/sharepoint`) — monorepo with library component + 5 web parts | webpart-builder | §2.1 | [x] |
| 1.0.2 | Configure `package.json` with all dependencies: PnPjs 3.x, Zustand 4.x, DevExtreme 22.2.x, devextreme-react, React Hook Form 7.x, @pnp/spfx-controls-react 3.x | webpart-builder | §2 | [x] |
| 1.0.3 | Add `spfx-toolkit` as a local dependency (`npm link` or file reference to `/Users/hemantmane/Development/spfx-toolkit`) | webpart-builder | §2.2 | [x] |
| 1.0.4 | Configure `tsconfig.json` — strict mode, path aliases for `sp-search-store` internal imports | webpart-builder | §2.1 | [x] |
| 1.0.5 | Configure `package-solution.json` — solution ID, library component reference, web part IDs | webpart-builder | §7.1 | [x] |
| 1.0.6 | Set up `.gitignore`, initial `README.md` | webpart-builder | — | [x] |
| 1.0.7 | Verify `gulp serve` launches workbench with empty web parts | webpart-builder | — | [x] |

**Exit criteria:** `gulp bundle --ship` succeeds. All 5 web parts + library component appear in workbench.

---

### Step 1.1 — TypeScript Interfaces (GATE)

> Every interface from §10.1 defined in `sp-search-store/interfaces/`. This is the contract the entire app builds on.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.1.1 | Create `ISearchStore` root store interface | store-architect | §10.1 | [x] |
| 1.1.2 | Create `IQuerySlice` — queryText, queryTemplate, scope, suggestions, abortController + action signatures | store-architect | §10.1 | [x] |
| 1.1.3 | Create `IFilterSlice` — activeFilters, availableRefiners, displayRefiners, filterConfig + action signatures | store-architect | §10.1 | [x] |
| 1.1.4 | Create `IResultSlice` — items, totalCount, currentPage, pageSize, sort, promotedResults + action signatures | store-architect | §10.1 | [x] |
| 1.1.5 | Create `IVerticalSlice` — currentVerticalKey, verticals, verticalCounts + action signatures | store-architect | §10.1 | [x] |
| 1.1.6 | Create `IUISlice` — activeLayoutKey, previewPanel, bulkSelection + action signatures | store-architect | §10.1 | [x] |
| 1.1.7 | Create `IUserSlice` — savedSearches, searchHistory, collections + action signatures | store-architect | §10.1 | [x] |
| 1.1.8 | Create `ISearchScope`, `ISuggestion`, `IActiveFilter`, `IRefiner`, `IRefinerValue` | store-architect | §10.1 | [x] |
| 1.1.9 | Create `ISearchResult`, `IPersonaInfo`, `ISortField` | store-architect | §10.1 | [x] |
| 1.1.10 | Create `ISearchDataProvider`, `ISearchQuery`, `ISearchResponse`, `IManagedProperty` | search-provider | §10.1 | [x] |
| 1.1.11 | Create `ISuggestionProvider`, `IActionProvider` | search-provider | §10.1 | [x] |
| 1.1.12 | Create `ILayoutDefinition`, `IFilterTypeDefinition` | layout-builder / filter-builder | §10.1 | [x] |
| 1.1.13 | Create `IFilterConfig`, `IFilterValueFormatter` | filter-builder | §10.1, §3.3.5 | [x] |
| 1.1.14 | Create `IVerticalDefinition` | store-architect | §10.1 | [x] |
| 1.1.15 | Create `ISavedSearch`, `ISearchCollection`, `ISearchHistoryEntry`, `IClickedItem` | search-manager | §10.1 | [x] |
| 1.1.16 | Create `IPromotedResult`, `IPromotedResultRule` | search-manager | §10.1 | [x] |
| 1.1.17 | Create `IRegistryContainer`, `Registry<T>` interface | store-architect | §10.1 | [x] |
| 1.1.18 | Create `interfaces/index.ts` barrel export | store-architect | — | [x] |

**Exit criteria:** All interfaces compile. `sp-search-store/interfaces/index.ts` exports everything. No circular dependencies.

---

### Step 1.2 — Registry Infrastructure (GATE)

> Generic `Registry<T>` class + all 5 typed registries. No built-in providers yet — just the container.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.2.1 | Implement generic `Registry<T extends { id: string }>` class — `register()`, `get()`, `getAll()`, `freeze()`. Duplicate IDs warn, first wins. `force: true` overrides. | store-architect | §4.4.6 | [x] |
| 1.2.2 | Create `DataProviderRegistry` extending `Registry<ISearchDataProvider>` | search-provider | §4.4.1 | [x] |
| 1.2.3 | Create `SuggestionProviderRegistry` extending `Registry<ISuggestionProvider>` | search-provider | §4.4.2 | [x] |
| 1.2.4 | Create `ActionProviderRegistry` extending `Registry<IActionProvider>` | search-provider | §4.4.3 | [x] |
| 1.2.5 | Create `LayoutRegistry` extending `Registry<ILayoutDefinition>` | layout-builder | §4.4.4 | [x] |
| 1.2.6 | Create `FilterTypeRegistry` extending `Registry<IFilterTypeDefinition>` | filter-builder | §4.4.5 | [x] |
| 1.2.7 | Write unit tests for `Registry<T>` — register, get, getAll, freeze, duplicate handling, force override | testing | — | [x] |

**Exit criteria:** All registries instantiate. Freeze prevents mutation. Tests pass.

---

### Step 1.3 — Zustand Store & Library Component (GATE)

> The backbone: store slices, store registry, library component shell.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.3.1 | Implement `querySlice` — setQueryText, setScope, setSuggestions, cancelSearch (aborts AbortController) | store-architect | §4.1.1 | [x] |
| 1.3.2 | Implement `filterSlice` — setRefiner, removeRefiner, clearAllFilters, setAvailableRefiners | store-architect | §4.1.1 | [x] |
| 1.3.3 | Implement `resultSlice` — setResults, setPage, setSort, promotedResults, isLoading, error | store-architect | §4.1.1 | [x] |
| 1.3.4 | Implement `verticalSlice` — setVertical, setVerticalCounts | store-architect | §4.1.1 | [x] |
| 1.3.5 | Implement `uiSlice` — setLayout, toggleSearchManager, setPreviewItem, toggleSelection | store-architect | §4.1.1 | [x] |
| 1.3.6 | Implement `userSlice` (stub) — empty arrays, placeholder action signatures for Phase 3 | store-architect | §4.1.1 | [x] |
| 1.3.7 | Implement `createStore(searchContextId)` — combines all slices + registries into one Zustand store | store-architect | §4.1 | [x] |
| 1.3.8 | Implement store registry — `getStore(id)` creates/returns from `Map<string, Store>`, `disposeStore(id)` cleans up | store-architect | §4.1 | [x] |
| 1.3.9 | Implement SPFx Library Component class (`SpSearchStoreLibrary`) — exposes `getStore()`, `disposeStore()` via SPFx library API | store-architect | §4.2 | [x] |
| 1.3.10 | Wire registries into store: `store.registries.dataProviders`, `.suggestions`, `.actions`, `.layouts`, `.filterTypes` | store-architect | §4.4.6 | [x] |
| 1.3.11 | Write unit tests for each slice — state mutations, action methods | testing | — | [x] |
| 1.3.12 | Write unit tests for store registry — create, retrieve, dispose, multi-instance isolation | testing | — | [x] |

**Exit criteria:** `getStore('test')` returns a working store. Slices mutate correctly. Two stores with different IDs are fully isolated.

---

### Step 1.4 — URL Sync Middleware

> Bi-directional URL ↔ store synchronization.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.4.1 | Implement URL serialization: store state → URL params (q, f, v, s, p, sc, l, sv) | store-architect | §4.1.2 | [x] |
| 1.4.2 | Implement URL deserialization: URL params → store state (page load restoration) | store-architect | §4.1.2 | [x] |
| 1.4.3 | Implement multi-context namespacing: `?ctx1.q=...&ctx2.q=...` when multiple stores exist | store-architect | §4.1.2 | [x] |
| 1.4.4 | Implement state version tag (`sv=1`) in serialized URLs | store-architect | §4.1.2 | [x] |
| 1.4.5 | Integrate middleware into `createStore()` — subscribe to state changes, push URL updates | store-architect | §4.1.2 | [x] |
| 1.4.6 | Handle browser back/forward (popstate) — restore store from URL | store-architect | §4.1.2 | [x] |
| 1.4.7 | Write unit tests for URL serialization round-trips (serialize → deserialize = original state) | testing | — | [x] |

**Exit criteria:** Changing query text updates URL. Pasting URL with `?q=test&v=documents` restores state. Back button works.

---

### Step 1.5 — Token Service & Search Service (GATE)

> Query construction engine. The single most reused piece of logic.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.5.1 | Implement `TokenService` — resolve `{searchTerms}`, `{Site.ID}`, `{Site.URL}`, `{Hub}`, `{Today}`, `{Today+N}`, `{User.*}`, `{PageContext.*}` | search-provider | §5.1, §10.5 | [x] |
| 1.5.2 | Implement KQL query assembly: template + user query + active filters → final KQL string | search-provider | §5.1 | [x] |
| 1.5.3 | Implement refinement token encoding: filter values → FQL/KQL refinement tokens per field type | search-provider | §3.3.4 | [x] |
| 1.5.4 | Implement sort parameter construction: `ISortField` → SharePoint `SortList` format | search-provider | §3.2.2 | [x] |
| 1.5.5 | Implement request coalescing: token resolution + query construction computed once, shared across results + count queries | search-provider | §4.3.3 | [x] |
| 1.5.6 | Implement `AbortController` lifecycle: create per search cycle, abort previous before new, pass signal to all API calls | search-provider | §4.3.3 | [x] |
| 1.5.7 | Write unit tests for TokenService — each token type with mock PageContext | testing | — | [x] |
| 1.5.8 | Write unit tests for KQL assembly — templates, filters, edge cases (empty query, no filters, etc.) | testing | — | [x] |

**Exit criteria:** `buildKqlQuery({ queryText: 'test', filters: [...] })` returns correct KQL. Token resolution handles all token types.

---

### Step 1.6 — SharePoint Search Provider (GATE)

> The default data provider — connects everything to the SharePoint Search API.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.6.1 | Implement `SharePointSearchProvider.execute()` — calls PnPjs `sp.search()` with constructed query, maps to `ISearchResponse` | search-provider | §3.2.1 | [x] |
| 1.6.2 | Implement result mapping: raw SharePoint search results → `ISearchResult[]` (title, url, summary, author, dates, fileType, fileSize, properties bag) | search-provider | §3.2.1 | [x] |
| 1.6.3 | Implement refiner parsing: raw refiner data → `IRefiner[]` with `IRefinerValue[]` (name, value token, count) | search-provider | §3.3 | [x] |
| 1.6.4 | Implement paging: `StartRow` calculation from `currentPage * pageSize` | search-provider | §3.2.1 | [x] |
| 1.6.5 | Implement `CollapseSpecification` support with sortability validation — warn + skip if property not sortable | search-provider | §3.2.6 | [x] |
| 1.6.6 | Implement `getSchema()` — fetch managed property metadata for Schema Helper | search-provider | §3.2.7 | [x] |
| 1.6.7 | Register `SharePointSearchProvider` as default in `DataProviderRegistry` during store creation | search-provider | §4.4.1 | [x] |
| 1.6.8 | Write unit tests with mocked PnPjs responses — result mapping, refiner parsing, error handling | testing | — | [x] |

**Exit criteria:** Provider takes an `ISearchQuery`, calls PnPjs, returns normalized `ISearchResponse`. Refiners are correctly parsed. CollapseSpecification validates sortability.

---

### Step 1.7 — Search Orchestrator

> The glue: listens to store changes, triggers search execution through the provider, dispatches results back.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.7.1 | Implement `SearchOrchestrator` — subscribes to querySlice + filterSlice + verticalSlice changes, triggers provider.execute() | search-provider | §5.1 | [x] |
| 1.7.2 | Wire AbortController: cancel previous search on every new trigger | search-provider | §4.3.3 | [x] |
| 1.7.3 | Dispatch results to `resultSlice.setResults()` and refiners to `filterSlice.setAvailableRefiners()` | search-provider | §5.1 | [x] |
| 1.7.4 | Implement parallel vertical count queries: `RowLimit=0` per vertical, shared AbortController, dispatch to `verticalSlice.setVerticalCounts()` | search-provider | §5.1 | [x] |
| 1.7.5 | Implement error handling: network errors, 4xx/5xx, aborted requests (don't show error for user-cancelled) | search-provider | §5.1 | [x] |
| 1.7.6 | Implement loading states: `resultSlice.isLoading = true` on search start, `false` on complete/error | search-provider | §5.1 | [x] |

**Exit criteria:** Changing `querySlice.queryText` automatically triggers a search and populates `resultSlice.items`. Vertical counts update in parallel. Previous searches are aborted.

---

### Step 1.8 — Search Box Web Part

> First visible UI. Query input → store → results.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.8.1 | Implement `SPSearchBoxWebPart.ts` — onInit (SPContext + getStore), render, property pane with `searchContextId` | webpart-builder | §3.1 | [x] |
| 1.8.2 | Implement `SearchBox.tsx` — text input with debounce (configurable, default 300ms), dispatches `setQueryText()` | webpart-builder | §3.1.1 | [x] |
| 1.8.3 | Implement search behavior: configurable "onEnter" / "onButton" / "both" | webpart-builder | §3.1.1 | [x] |
| 1.8.4 | Implement clear/reset button — clears query, filters, verticals, sort to defaults | webpart-builder | §3.1.1 | [x] |
| 1.8.5 | Implement `ScopeSelector.tsx` — dropdown with configurable scopes (Current Site, Hub, All SharePoint, custom), dispatches `setScope()` | webpart-builder | §3.1.1 | [x] |
| 1.8.6 | Implement property pane: placeholder, debounceMs, searchBehavior, enableScopeSelector, searchScopes, enableSuggestions, enableSearchManager | webpart-builder | §3.1.2 | [x] |
| 1.8.7 | Wrap root component with `ErrorBoundary` from spfx-toolkit | webpart-builder | §6.1 | [x] |
| 1.8.8 | Fix duplicate orchestrator — use shared orchestrator via `initializeSearchContext()` from registry | webpart-builder | §4.1 | [x] |

**Exit criteria:** Typing in the search box updates the Zustand store after debounce. Scope selector changes scope. Clear resets state.

---

### Step 1.9 — Search Results Web Part (List + Compact Layouts)

> Core result display with two basic layouts.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.9.1 | Implement `SPSearchResultsWebPart.ts` — onInit, render, property pane with `searchContextId`, `queryTemplate`, `selectedProperties`, `pageSize`, `defaultLayout` | webpart-builder | §3.2.5 | [x] |
| 1.9.2 | Implement `SearchResults.tsx` — subscribes to `resultSlice`, renders active layout component, loading shimmer, empty state, error state | webpart-builder | §3.2 | [x] |
| 1.9.3 | Implement `ResultToolbar.tsx` — result count display, layout switcher toggle, sort dropdown | webpart-builder | §3.2 | [x] |
| 1.9.4 | Implement `ListLayout.tsx` — Google-style result cards per §3.2.2 spec (icon + title + URL breadcrumb + excerpt + metadata line) | layout-builder | §3.2.2 | [x] |
| 1.9.5 | Implement hit-highlighting in list layout — `<mark>` tags with theme-aware styling, sanitize SharePoint `<ddd>` tags | layout-builder | §3.2.2 | [x] |
| 1.9.6 | Implement `CompactLayout.tsx` — single-line-per-result (icon + title + author + modified + type), hover excerpt | layout-builder | §3.2.2 | [x] |
| 1.9.7 | Implement numbered pagination component — page numbers, next/prev, syncs with `resultSlice.currentPage` | layout-builder | §3.2.1 | [x] |
| 1.9.8 | Register ListLayout and CompactLayout in LayoutRegistry with proper `ILayoutDefinition` | layout-builder | §4.4.4 | [x] |
| 1.9.9 | Implement basic result selection — checkbox on hover, `uiSlice.toggleSelection()` | layout-builder | §3.2.2 | [x] |
| 1.9.10 | Implement keyboard navigation — arrow keys between results, Enter to select, Space to toggle selection | layout-builder | §3.2.2 | [x] |
| 1.9.11 | Wrap root with `ErrorBoundary`, implement loading shimmer via Fluent UI `Shimmer` | webpart-builder | §6.1 | [x] |

**Exit criteria:** Search results display in List layout (default) and Compact layout. Layout switcher toggles between them. Pagination works. Hit-highlighting visible.

---

### Step 1.10 — Search Filters Web Part (Checkbox + Date Range)

> Basic refinement: two filter types that cover 80% of use cases.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.10.1 | Implement `SPSearchFiltersWebPart.ts` — onInit, render, property pane with `searchContextId`, `filters[]`, `applyMode`, `operatorBetweenFilters` | webpart-builder | §3.3.7 | [x] |
| 1.10.2 | Implement `SearchFilters.tsx` — reads `filterSlice.availableRefiners` + `filterConfig`, renders filter groups | webpart-builder | §3.3 | [x] |
| 1.10.3 | Implement `FilterGroup.tsx` — spfx-toolkit Card accordion per filter, expand/collapse with persistence via `useLocalStorage` | filter-builder | §3.3.2 | [x] |
| 1.10.4 | Implement `CheckboxFilter.tsx` — multi-select checkboxes with result counts, "search within" text filter, show more/less | filter-builder | §3.3.1 | [x] |
| 1.10.5 | Implement `DateRangeFilter.tsx` — preset buttons (Today, This Week, This Month, This Year, Custom) + DevExtreme DateRangeBox | filter-builder | §3.3.1, §3.3.4A | [x] |
| 1.10.6 | Implement `DateFilterFormatter` — FQL `range(datetime("..."), datetime("..."))` generation, timezone handling (UTC for FQL, local for display) | filter-builder | §3.3.4A, §3.3.5 | [x] |
| 1.10.7 | Implement `DefaultFilterFormatter` — pass-through for checkbox string values | filter-builder | §3.3.5 | [x] |
| 1.10.8 | Implement instant apply mode: filter change → `setRefiner()` → triggers search | filter-builder | §3.3.2 | [x] |
| 1.10.9 | Implement manual apply mode: stage changes, Apply button dispatches `applyFilters()` | filter-builder | §3.3.2 | [x] |
| 1.10.10 | Implement individual clear + global "Clear All" button | filter-builder | §3.3.2 | [x] |
| 1.10.11 | Register CheckboxFilter and DateRangeFilter in FilterTypeRegistry | filter-builder | §4.4.5 | [x] |
| 1.10.12 | Wrap root with `ErrorBoundary` | webpart-builder | §6.1 | [x] |

**Exit criteria:** Checkbox filter shows refiner values with counts. Date range filter generates FQL range tokens. Selecting a filter updates results. Clear works.

---

### Step 1.11 — Search Verticals Web Part

> Tab navigation with badge counts.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.11.1 | Implement `SPSearchVerticalsWebPart.ts` — onInit, render, property pane with `searchContextId`, `verticals[]`, `showCounts`, `tabStyle` | webpart-builder | §3.4.2 | [x] |
| 1.11.2 | Implement `SearchVerticals.tsx` — horizontal tab bar reading from `verticalSlice`, dispatches `setVertical()` on tab click | webpart-builder | §3.4.1 | [x] |
| 1.11.3 | Implement `VerticalTab.tsx` — tab label + icon + badge count, active state styling | webpart-builder | §3.4.1 | [x] |
| 1.11.4 | Implement badge counts display from `verticalSlice.verticalCounts` | webpart-builder | §3.4.1 | [x] |
| 1.11.5 | Implement `hideEmptyVerticals` option — dimmed or hidden tabs with zero counts | webpart-builder | §3.4.1 | [x] |
| 1.11.6 | Implement overflow handling — excess tabs collapse into "More" dropdown on narrow screens | webpart-builder | §3.4.1 | [x] |
| 1.11.7 | Implement tab styles: "tabs", "pills", "underline" (configurable) | webpart-builder | §3.4.2 | [x] |
| 1.11.8 | Wrap root with `ErrorBoundary` | webpart-builder | §6.1 | [x] |

**Exit criteria:** Tabs display with counts. Clicking a tab switches vertical, triggers new search with vertical-specific query template/result source. Empty verticals handled.

---

### Step 1.12 — Hidden List Provisioning Script

> PowerShell script that creates the 4 hidden lists with proper columns and indexes.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.12.1 | Create `Provision-SPSearchLists.ps1` — idempotent, checks existence before creating | search-manager | §7.2 | [x] |
| 1.12.2 | Provision `SearchSavedQueries` list — columns per §3.5.6, Hidden=true, indexes on Title, EntryType, Category, SharedWith, LastUsed | search-manager | §3.5.6 | [x] |
| 1.12.3 | Provision `SearchHistory` list — columns per §3.5.6, Hidden=true, indexes on Author, SearchTimestamp, QueryHash, Vertical | search-manager | §3.5.6 | [x] |
| 1.12.4 | Provision `SearchCollections` list — columns per §3.5.6, Hidden=true, indexes on Title, ItemUrl, CollectionName, SharedWith | search-manager | §3.5.6 | [x] |
| 1.12.5 | Provision `SearchConfiguration` list — columns per §3.5.6, Hidden=true, indexes on Title, ConfigType, IsActive, ExpiresAt | search-manager | §3.5.6 | [x] |
| 1.12.6 | Set list permissions: SearchHistory = Add + Edit Own, SearchConfiguration = Admin-only write, others = Add Items | search-manager | §8.2 | [x] |
| 1.12.7 | Seed default SearchConfiguration entries: default scopes, layout mappings | search-manager | §7.2 | [x] |
| 1.12.8 | Verify index creation succeeds before marking provisioning complete | search-manager | §3.5.6 | [x] |

**Exit criteria:** Script runs idempotently. All 4 lists created with correct columns, indexes, and permissions. Re-running is safe.

---

### Step 1.13 — Phase 1 Integration Testing

> End-to-end: type query → see results → apply filter → switch vertical → URL updates.

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 1.13.1 | Test end-to-end: Search Box → store → SearchOrchestrator → SharePointSearchProvider → resultSlice → Search Results | testing | §5.1 | [ ] |
| 1.13.2 | Test filter flow: select checkbox filter → filterSlice → re-execute search → new results + updated refiner counts | testing | §5.2 | [ ] |
| 1.13.3 | Test vertical switch: click tab → verticalSlice → re-execute with vertical-specific query → results + counts update | testing | §5.1 | [ ] |
| 1.13.4 | Test URL sync: changing state updates URL, loading URL restores state, back button works | testing | §4.1.2 | [ ] |
| 1.13.5 | Test multi-instance: two sets of web parts with different `searchContextId` — fully isolated | testing | §4.1 | [ ] |
| 1.13.6 | Test abort: rapid typing cancels previous requests, no stale results displayed | testing | §4.3.3 | [ ] |
| 1.13.7 | Verify `gulp bundle --ship` bundle size — spfx-toolkit and Fluent UI are tree-shaken | testing | §4.3.1 | [x] |

**Exit criteria:** Full Phase 1 functionality works in SPFx workbench. No console errors. Bundle size reasonable.

---

## Phase 2: Rich Layouts

> **Goal:** All 6 layouts working + Result Detail Panel + layout switcher. The search experience goes from functional to impressive.

---

### Step 2.1 — DataGrid Layout (DevExtreme)

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.1.1 | Implement `DataGridLayout.tsx` — lazy-loaded via `React.lazy()`, virtual scrolling, column config from selected properties | layout-builder | §3.2.2 | [x] |
| 2.1.2 | Implement 12 type-aware cell renderers: Title, Persona, Date, FileSize, FileType, URL, Taxonomy, Boolean, Number, Tags, Thumbnail, Text | layout-builder | §3.2.2 | [x] |
| 2.1.3 | Implement server-side sorting — column header click dispatches `setSort()`, triggers API call with `SortList` | layout-builder | §3.2.2 | [x] |
| 2.1.4 | Implement client-side secondary sort — Shift+Click on second column | layout-builder | §3.2.2 | [x] |
| 2.1.5 | Implement client-side column filtering with type-aware filter editors (text contains, date range, person picker, file type multi-select, number range, boolean toggle, taxonomy dropdown) | layout-builder | §3.2.2 | [x] |
| 2.1.6 | Implement column grouping with drag-to-group area | layout-builder | §3.2.2 | [x] |
| 2.1.7 | Implement column reordering, resizing, visibility toggle (column chooser) | layout-builder | §3.2.2 | [x] |
| 2.1.8 | Implement row selection: single + multi-select with Shift+Click range | layout-builder | §3.2.2 | [x] |
| 2.1.9 | Implement export to Excel and CSV via DevExtreme `exportDataGrid` | layout-builder | §3.2.2 | [x] |
| 2.1.10 | Implement master-detail row expansion (inline document preview) | layout-builder | §3.2.2 | [x] |
| 2.1.11 | Implement responsive behavior: auto-switch to card mode on < 768px | layout-builder | §3.2.2 | [x] |
| 2.1.12 | Register DataGridLayout in LayoutRegistry | layout-builder | §4.4.4 | [x] |

---

### Step 2.2 — Card Layout (spfx-toolkit)

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.2.1 | Implement `CardLayout.tsx` — spfx-toolkit Card per result, configurable metadata fields | layout-builder | §3.2.2 | [x] |
| 2.2.2 | Implement Card header: document title + icon + quick actions | layout-builder | §3.2.2 | [x] |
| 2.2.3 | Implement accordion pattern for grouping cards by property (content type, site, custom) | layout-builder | §3.2.2 | [x] |
| 2.2.4 | Implement card maximize — expand single result into full detail view | layout-builder | §3.2.2 | [x] |
| 2.2.5 | Implement lazy loading of card content when scrolled into view | layout-builder | §3.2.2 | [x] |
| 2.2.6 | Implement responsive grid: 1 col mobile, 2 tablet, 3-4 desktop | layout-builder | §3.2.2 | [x] |
| 2.2.7 | Register CardLayout in LayoutRegistry | layout-builder | §4.4.4 | [x] |

---

### Step 2.3 — People Layout

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.3.1 | Implement `PeopleLayout.tsx` — Fluent UI Persona + spfx-toolkit UserPersona per result | layout-builder | §3.2.2 | [x] |
| 2.3.2 | Display: photo, name, job title, department, office, presence indicator | layout-builder | §3.2.2 | [x] |
| 2.3.3 | Implement contact actions: email, Teams chat deep link, call | layout-builder | §3.2.2 | [x] |
| 2.3.4 | Implement org chart info: direct reports count, manager name | layout-builder | §3.2.2 | [x] |
| 2.3.5 | Implement expandable "Recent documents by this person" section | layout-builder | §3.2.2 | [x] |
| 2.3.6 | Implement PersonaCard on hover | layout-builder | §3.2.2 | [x] |
| 2.3.7 | Register PeopleLayout in LayoutRegistry | layout-builder | §4.4.4 | [x] |

---

### Step 2.4 — Document Gallery Layout

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.4.1 | Implement `DocumentGalleryLayout.tsx` — thumbnail grid using SharePoint preview thumbnails API | layout-builder | §3.2.2 | [x] |
| 2.4.2 | Implement configurable thumbnail sizes: small (120px), medium (200px), large (300px) | layout-builder | §3.2.2 | [x] |
| 2.4.3 | Implement hover overlay: title, file type, modified date | layout-builder | §3.2.2 | [x] |
| 2.4.4 | Implement lightbox view for images | layout-builder | §3.2.2 | [x] |
| 2.4.5 | Implement masonry vs fixed grid option | layout-builder | §3.2.2 | [x] |
| 2.4.6 | Implement infinite scroll pagination | layout-builder | §3.2.2 | [x] |
| 2.4.7 | Register DocumentGalleryLayout in LayoutRegistry | layout-builder | §4.4.4 | [x] |

---

### Step 2.5 — Result Detail Panel

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.5.1 | Implement `ResultDetailPanel.tsx` — Fluent UI Panel, subscribes to `uiSlice.previewPanel`, lazy-loaded on first open | layout-builder | §3.2.3 | [x] |
| 2.5.2 | Implement `DocumentPreview.tsx` — WOPI frame for Office docs/PDFs, inline image, video player, fallback for unsupported types | layout-builder | §3.2.3A | [x] |
| 2.5.3 | Implement WOPI frame URL construction: `{siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={uniqueId}&action=interactivepreview` | layout-builder | §3.2.3A | [x] |
| 2.5.4 | Implement preview loading state (Shimmer) and error state (file icon + "Preview unavailable") | layout-builder | §3.2.3A | [x] |
| 2.5.5 | Implement `MetadataDisplay.tsx` — type-aware formatted values using same renderer patterns as DataGrid cells | layout-builder | §3.2.3B | [x] |
| 2.5.6 | Integrate spfx-toolkit `LazyVersionHistory` — last 5 versions, "View All" loads full history, click version to preview | layout-builder | §3.2.3C | [x] |
| 2.5.7 | Implement `RelatedDocuments.tsx` — lazy-loaded, same library + similar metadata queries, max 5 results | layout-builder | §3.2.3D | [x] |
| 2.5.8 | Implement quick actions toolbar: Open (split button), Download, Copy Link, Share (stub), Pin (stub), View in Library | layout-builder | §3.2.3E | [x] |
| 2.5.9 | Wire click-on-result → `uiSlice.setPreviewItem()` → panel opens in all layouts | layout-builder | §3.2.3 | [x] |

---

### Step 2.6 — Layout Switcher & Refiner Stability

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.6.1 | Implement layout switcher in `ResultToolbar.tsx` — icon buttons for each enabled layout, dispatches `uiSlice.setLayout()` | webpart-builder | §3.2.2 | [x] |
| 2.6.2 | Implement refiner stability mode in `filterSlice` — `displayRefiners` debounced from `availableRefiners` (configurable, default 500ms) | store-architect | §4.3.3 | [x] |
| 2.6.3 | Update SearchFilters to render from `displayRefiners` instead of `availableRefiners` when stability mode is on | filter-builder | §4.3.3 | [x] |

---

### Step 2.7 — Phase 2 Integration Testing

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 2.7.1 | Test all 6 layouts render correctly with search results | testing | §3.2.2 | [ ] |
| 2.7.2 | Test layout switcher: toggle between all layouts preserves results, selection, scroll position | testing | §3.2.2 | [ ] |
| 2.7.3 | Test DataGrid: sort, filter, group, export, virtual scrolling, responsive card mode | testing | §3.2.2 | [ ] |
| 2.7.4 | Test Result Detail Panel: opens on click, preview loads, metadata formatted, version history works | testing | §3.2.3 | [ ] |
| 2.7.5 | Test refiner stability: rapid typing doesn't cause filter options to flicker | testing | §4.3.3 | [ ] |
| 2.7.6 | Verify DataGrid lazy loading — bundle chunk only loaded when grid layout selected | testing | §4.3.1 | [ ] |

---

## Phase 3: User Features

> **Goal:** Search Manager fully operational — saved searches, sharing, collections, history, promoted results. The features that differentiate SP Search from PnP.

---

### Step 3.1 — Search Manager Service (CRUD Layer)

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.1.1 | Implement `SearchManagerService` — CRUD for SearchSavedQueries list (create, read, update, delete) | search-manager | §3.5.1 | [x] |
| 3.1.2 | Implement search state serialization: full Zustand state → JSON for `SearchState` column | search-manager | §5.3 | [x] |
| 3.1.3 | Implement search state deserialization: `SearchState` JSON → restore full store state | search-manager | §5.3 | [x] |
| 3.1.4 | Implement CRUD for SearchCollections list | search-manager | §3.5.3 | [x] |
| 3.1.5 | Implement CRUD for SearchHistory list — ALL queries filter by `Author eq [Me]` FIRST | search-manager | §3.5.4 | [x] |
| 3.1.6 | Implement history deduplication via QueryHash (SHA-256 of full query state) | search-manager | §3.5.4 | [x] |
| 3.1.7 | Implement CRUD for SearchConfiguration list (admin operations) | search-manager | §3.5.6 | [x] |
| 3.1.8 | Wire `SearchManagerService` into `userSlice` actions: `saveSearch()`, `loadHistory()`, `addToHistory()` | search-manager | §4.1.1 | [x] |

---

### Step 3.2 — Item-Level Permissions

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.2.1 | Implement `breakRoleInheritance()` on saved search creation — author gets Full Control only | search-manager | §8.2 | [x] |
| 3.2.2 | Implement `addRoleAssignment()` on share — each SharedWith user gets Read permission | search-manager | §8.2 | [x] |
| 3.2.3 | Implement permission update when SharedWith changes — add/remove role assignments | search-manager | §8.2 | [x] |
| 3.2.4 | Apply same pattern to SearchCollections sharing | search-manager | §8.2 | [x] |

---

### Step 3.3 — Search Manager Web Part

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.3.1 | Implement `SPSearchManagerWebPart.ts` — standalone web part mode | webpart-builder | §3.5 | [x] |
| 3.3.2 | Implement `SearchManager.tsx` — root component with tab navigation: Saved, Shared, Collections, History | webpart-builder | §3.5 | [x] |
| 3.3.3 | Implement `SavedSearchList.tsx` — list of saved searches with name, query preview, date, result count. Load/edit/delete actions | search-manager | §3.5.1 | [x] |
| 3.3.4 | Implement Save Search dialog — name input, category picker, save current state | search-manager | §3.5.1 | [x] |
| 3.3.5 | Implement "Shared With Me" section — searches shared by others | search-manager | §3.5.2 | [x] |
| 3.3.6 | Implement `SearchCollections.tsx` — collection list, pin/unpin, view collection items, manage (rename, delete, reorder, merge) | search-manager | §3.5.3 | [x] |
| 3.3.7 | Implement `SearchHistory.tsx` — chronological list, re-execute any past search, clear all | search-manager | §3.5.4 | [x] |
| 3.3.8 | Implement Search Manager as Fluent UI Panel (triggered from Search Box icon button) | webpart-builder | §3.5 | [x] |
| 3.3.9 | Lazy-load panel content on first open | webpart-builder | §4.3.1 | [x] |

---

### Step 3.4 — Search Sharing

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.4.1 | Implement `ShareSearchDialog.tsx` — tabbed dialog: URL, Email, Teams, Users | search-manager | §3.5.2 | [x] |
| 3.4.2 | Share via URL: encode full state in URL params, copy to clipboard with toast confirmation | search-manager | §3.5.2 | [x] |
| 3.4.3 | Share via Email: `mailto:` with search description + link + optional top N results | search-manager | §3.5.2 | [x] |
| 3.4.4 | Share to Teams: deep link `https://teams.microsoft.com/l/chat/0/0?message={encoded}` | search-manager | §3.5.2 | [x] |
| 3.4.5 | Share to Users: PnP PeoplePicker for user selection, create SharedSearch entry, set item-level permissions | search-manager | §3.5.2 | [x] |

---

### Step 3.5 — Search History Auto-Logging

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.5.1 | Implement auto-logging in SearchOrchestrator — after successful search, dispatch `addToHistory()` asynchronously | search-manager | §5.1 | [x] |
| 3.5.2 | Implement clicked item tracking: when user clicks a result, log `{ url, title, position, timestamp }` to history entry | search-manager | §3.5.4 | [x] |
| 3.5.3 | Implement configurable history cleanup TTL (30/60/90 days) in SearchConfiguration | search-manager | §3.5.4 | [x] |

---

### Step 3.6 — StateId Deep Link Fallback

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.6.1 | Detect when serialized URL exceeds 2,000 chars | store-architect | §4.1.2 | [x] |
| 3.6.2 | Save full state JSON to SearchConfiguration list with `ConfigType: StateSnapshot`, `ExpiresAt` TTL | store-architect | §4.1.2 | [x] |
| 3.6.3 | Replace URL with `?sid=<itemId>` — on page load, detect `sid`, fetch state from list, restore | store-architect | §4.1.2 | [x] |

---

### Step 3.7 — Promoted Results / Best Bets

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.7.1 | Implement promoted result rule evaluation: match query against `contains`, `equals`, `regex`, `kql` rules from SearchConfiguration | search-manager | §3.6.1 | [x] |
| 3.7.2 | Implement `PromotedResultsBlock.tsx` — "Recommended" block above organic results, visually distinct styling | layout-builder | §3.6.2 | [x] |
| 3.7.3 | Implement layout-adaptive rendering: card style in Card Layout, row in DataGrid/List/Compact | layout-builder | §3.6.2 | [x] |
| 3.7.4 | Implement dismissible promoted results (session-only, stored in uiSlice) | layout-builder | §3.6.2 | [x] |
| 3.7.5 | Implement configurable max promoted results per query (default 3) | search-manager | §3.6.4 | [x] |

---

### Step 3.8 — Active Filter Pill Bar

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.8.1 | Implement `ActiveFilterPillBar.tsx` — horizontal strip of dismissible pills, rendered by Search Results web part | filter-builder | §3.3.3 | [x] |
| 3.8.2 | Implement pill rendering: `{Filter Name}: {Human-Readable Value} x` per filter | filter-builder | §3.3.3 | [x] |
| 3.8.3 | Multi-value filters combined into ONE pill with comma-separated values | filter-builder | §3.3.3 | [x] |
| 3.8.4 | Pill click dismisses filter via `removeRefiner()`, re-executes search | filter-builder | §3.3.3 | [x] |
| 3.8.5 | "Clear All" link at end dispatches `clearAllFilters()` | filter-builder | §3.3.3 | [x] |
| 3.8.6 | Human-readable display via `IFilterValueFormatter` for each field type | filter-builder | §3.3.3 | [x] |
| 3.8.7 | Animate pill add/remove (Fluent UI motion tokens) | filter-builder | §3.3.3 | [x] |
| 3.8.8 | Sticky behavior when filter panel is in sidebar layout | filter-builder | §3.3.3 | [x] |

---

### Step 3.9 — Recent Searches Suggestion Provider

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.9.1 | Implement `RecentSearchProvider` — queries SearchHistory list for current user's recent searches | search-provider | §4.4.2 | [x] |
| 3.9.2 | Implement `SuggestionDropdown.tsx` in Search Box — dropdown below input showing grouped suggestions | webpart-builder | §3.1.1 | [x] |
| 3.9.3 | Register RecentSearchProvider in SuggestionProviderRegistry | search-provider | §4.4.2 | [x] |

---

### Step 3.10 — Phase 3 Integration Testing

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 3.10.1 | Test save/load/delete search lifecycle | testing | §3.5.1 | [ ] |
| 3.10.2 | Test sharing: URL, email, Teams, user-specific with item-level permissions | testing | §3.5.2 | [ ] |
| 3.10.3 | Test collections: create, pin from result, view, share, manage | testing | §3.5.3 | [ ] |
| 3.10.4 | Test history: auto-log, deduplication, cleanup, re-execute | testing | §3.5.4 | [ ] |
| 3.10.5 | Test promoted results: rule matching, display, dismissal | testing | §3.6 | [ ] |
| 3.10.6 | Test StateId fallback: long URL → `?sid=` → state restoration | testing | §4.1.2 | [ ] |
| 3.10.7 | Test pill bar: display, dismiss, clear all, human-readable formatting | testing | §3.3.3 | [ ] |

---

## Phase 4: Power Features

> **Goal:** Advanced query capabilities, remaining filter types, bulk actions, smart suggestions, audience targeting.

---

### Step 4.1 — Advanced Filter Types

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.1.1 | Implement `TaxonomyTreeFilter.tsx` — DevExtreme TreeView, hierarchical expand/collapse, multi-select, search-within | filter-builder | §3.3.1, §3.3.4C | [x] |
| 4.1.2 | Implement `TaxonomyFilterFormatter` — GP0|#GUID to term label resolution via PnP Taxonomy API, caching | filter-builder | §3.3.4C, §3.3.5 | [x] |
| 4.1.3 | Implement `PeoplePickerFilter.tsx` — PnP PeoplePicker, type-ahead against AAD, multi-select | filter-builder | §3.3.1, §3.3.4B | [x] |
| 4.1.4 | Implement `PeopleFilterFormatter` — claim string resolution to display names, cached profiles | filter-builder | §3.3.4B, §3.3.5 | [x] |
| 4.1.5 | Implement `SliderFilter.tsx` — DevExtreme RangeSlider, configurable min/max/step, file size formatting | filter-builder | §3.3.1, §3.3.4E | [x] |
| 4.1.6 | Implement `NumericFilterFormatter` — FQL range(decimal(), decimal()), file size KB/MB/GB, currency | filter-builder | §3.3.4E, §3.3.5 | [x] |
| 4.1.7 | Implement `TagBoxFilter.tsx` — DevExtreme TagBox, tag-style multi-select with search | filter-builder | §3.3.1 | [x] |
| 4.1.8 | Implement `ToggleFilter.tsx` — Fluent UI Toggle, three-state (All/Yes/No) | filter-builder | §3.3.1, §3.3.4F | [x] |
| 4.1.9 | Implement `BooleanFilterFormatter` — "0"/"1" to Yes/No or custom labels | filter-builder | §3.3.4F, §3.3.5 | [x] |
| 4.1.10 | Register all new filter types in FilterTypeRegistry | filter-builder | §4.4.5 | [x] |

---

### Step 4.2 — Visual Query Builder

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.2.1 | Implement `QueryBuilder.tsx` — expandable panel below search box, DevExtreme-inspired filter builder UI | webpart-builder | §3.1.1 | [x] |
| 4.2.2 | Property dropdowns populated from available managed properties (Schema Helper data) | webpart-builder | §3.1.1 | [x] |
| 4.2.3 | Operator selection per property type (contains, equals, greater than, range, etc.) | webpart-builder | §3.1.1 | [x] |
| 4.2.4 | Value pickers matched to property type (text, date, number, person, taxonomy) | webpart-builder | §3.1.1 | [x] |
| 4.2.5 | Convert builder expression to KQL and dispatch to querySlice | webpart-builder | §3.1.1 | [x] |

---

### Step 4.3 — Visual Filter Builder

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.3.1 | Implement `VisualFilterBuilder.tsx` — DevExtreme FilterBuilder-inspired UI in Search Filters web part | filter-builder | §3.3.6 | [x] |
| 4.3.2 | AND/OR grouping of filter expressions | filter-builder | §3.3.6 | [x] |
| 4.3.3 | Convert builder expression to KQL refinement queries | filter-builder | §3.3.6 | [x] |

---

### Step 4.4 — Bulk Actions Toolbar

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.4.1 | Implement `BulkActionsToolbar.tsx` — appears when `uiSlice.bulkSelection.length > 0` | webpart-builder | §3.2.4 | [x] |
| 4.4.2 | Implement built-in action providers: `OpenAction`, `PreviewAction`, `ShareAction`, `PinAction`, `CopyLinkAction`, `DownloadAction` | search-provider | §4.4.3 | [x] |
| 4.4.3 | Register all built-in actions in ActionProviderRegistry | search-provider | §4.4.3 | [x] |
| 4.4.4 | Implement bulk share: share multiple items via URL/email/Teams | search-manager | §3.2.4 | [x] |
| 4.4.5 | Implement bulk download: download selected files individually | search-manager | §3.2.4 | [x] |
| 4.4.6 | Implement bulk pin: pin selected items to a collection | search-manager | §3.2.4 | [x] |
| 4.4.7 | Implement metadata comparison: side-by-side compare of 2-3 selected items | layout-builder | §3.2.4 | [x] |
| 4.4.8 | Implement export selected to Excel/CSV | layout-builder | §3.2.4 | [x] |

---

### Step 4.5 — Smart Suggestions

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.5.1 | Implement `TrendingQueryProvider` — popular queries aggregated from SearchHistory across org | search-provider | §4.4.2 | [x] |
| 4.5.2 | Implement `ManagedPropertyProvider` — suggests property values matching input (e.g., "Author: John Doe") | search-provider | §4.4.2 | [x] |
| 4.5.3 | Update `SuggestionDropdown.tsx` — merge suggestions from all providers, grouped sections, ranked by priority | webpart-builder | §3.1.1 | [x] |
| 4.5.4 | Register TrendingQueryProvider and ManagedPropertyProvider in SuggestionProviderRegistry | search-provider | §4.4.2 | [x] |

---

### Step 4.6 — Result Annotations / Tags

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.6.1 | Implement `ResultAnnotations.tsx` — tag input UI for personal tags ("Reviewed", "Important", "Follow-up") | search-manager | §3.5.5 | [x] |
| 4.6.2 | Store tags in SearchCollections list (Tags column) | search-manager | §3.5.5 | [x] |
| 4.6.3 | Implement shared tags option — visible to team members | search-manager | §3.5.5 | [x] |
| 4.6.4 | Display tag badges/labels in search result layouts | layout-builder | §3.5.5 | [x] |
| 4.6.5 | Tag-based filtering in Search Manager | search-manager | §3.5.5 | [x] |

---

### Step 4.7 — Audience Targeting

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.7.1 | Implement audience group checking — resolve current user's Azure AD security group memberships | search-provider | §3.4.1 | [x] |
| 4.7.2 | Apply audience targeting to verticals — hide tabs not targeted to current user | webpart-builder | §3.4.1 | [x] |
| 4.7.3 | Apply audience targeting to promoted results — only show rules matching user's groups | search-manager | §3.6.1 | [x] |

---

### Step 4.8 — Schema Helper (Property Pane Control)

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.8.1 | Implement `PropertyPaneSchemaHelper.ts` — "Browse Schema" button in property pane | webpart-builder | §3.2.7 | [x] |
| 4.8.2 | Fetch search schema via Search Administration API — searchable/filterable list of managed properties | search-provider | §3.2.7 | [x] |
| 4.8.3 | Display: property name, alias, type, queryable/retrievable/refinable/sortable flags | webpart-builder | §3.2.7 | [x] |
| 4.8.4 | Click to insert property into config field | webpart-builder | §3.2.7 | [x] |
| 4.8.5 | Permission check: fall back to text input if user lacks Search Admin permissions | webpart-builder | §3.2.7 | [x] |
| 4.8.6 | Cache schema in sessionStorage | search-provider | §3.2.7 | [x] |

---

### Step 4.9 — Phase 4 Integration Testing

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 4.9.1 | Test all 7 filter types with real SharePoint refiners | testing | §3.3.1 | [x] |
| 4.9.2 | Test taxonomy tree filter: hierarchical selection, parent includes children | testing | §3.3.4C | [x] |
| 4.9.3 | Test people picker filter: claim string resolution, caching | testing | §3.3.4B | [x] |
| 4.9.4 | Test visual query builder: expression to KQL conversion | testing | §3.1.1 | [x] |
| 4.9.5 | Test bulk actions: share, download, pin, compare, export | testing | §3.2.4 | [x] |
| 4.9.6 | Test audience targeting: verticals and promoted results visibility | testing | §3.4.1, §3.6.1 | [x] |
| 4.9.7 | Test smart suggestions: recent, trending, property value suggestions | testing | §4.4.2 | [x] |

---

## Phase 5: Polish & Optimization

> **Goal:** Production-ready. Optimized bundles, accessible, responsive, documented.

---

### Step 5.1 — Bundle Optimization

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.1.1 | Run `gulp bundle --ship --analyze-bundle` — identify all large chunks | testing | §4.3.1 | [x] |
| 5.1.2 | Verify spfx-toolkit tree-shaking — no barrel imports pulling full library | testing | §4.3.1 | [x] |
| 5.1.3 | Verify Fluent UI tree-shaking — no `@fluentui/react` root imports | testing | §4.3.1 | [x] |
| 5.1.4 | Verify DevExtreme lazy loading — DataGrid chunk only on grid layout | testing | §4.3.1 | [x] |
| 5.1.5 | Verify Detail Panel lazy loading — chunk only on first panel open | testing | §4.3.1 | [x] |
| 5.1.6 | Verify Search Manager panel lazy loading | testing | §4.3.1 | [x] |
| 5.1.7 | Profile each layout chunk size — identify optimization opportunities | testing | §4.3.1 | [x] |
| 5.1.8 | Reduce total .sppkg size to acceptable range for site-level app catalog | testing | §7.1 | [x] |

---

### Step 5.2 — Accessibility Audit (WCAG 2.1 AA)

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.2.1 | Keyboard navigation: all interactive elements reachable via Tab, operable via Enter/Space | webpart-builder | §10.3 | [x] |
| 5.2.2 | Screen reader: all web parts announce state changes, results, filter updates | webpart-builder | §10.3 | [x] |
| 5.2.3 | Focus management: focus moves logically between web parts, traps in modals/panels | webpart-builder | §10.3 | [x] |
| 5.2.4 | Color contrast: all text meets 4.5:1 ratio (use Fluent UI semantic tokens) | webpart-builder | §10.3 | [x] |
| 5.2.5 | ARIA labels: all buttons, inputs, regions properly labeled | webpart-builder | §10.3 | [x] |
| 5.2.6 | Live regions: search result count, filter changes announced to screen readers | webpart-builder | §10.3 | [x] |

---

### Step 5.3 — Responsive Design

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.3.1 | Test all layouts at mobile (320px), tablet (768px), desktop (1024px+) breakpoints | layout-builder | §10.3 | [ ] |
| 5.3.2 | Search Box: full width on mobile, inline scope selector collapse | webpart-builder | — | [ ] |
| 5.3.3 | Filters: collapse to panel/drawer on mobile | filter-builder | — | [ ] |
| 5.3.4 | Verticals: overflow to "More" dropdown on narrow screens | webpart-builder | §3.4.1 | [ ] |
| 5.3.5 | Detail Panel: full-screen on mobile | layout-builder | §3.2.3 | [ ] |

---

### Step 5.4 — Error Handling & Empty States

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.4.1 | Implement empty state for no results: helpful message + suggestions (check spelling, broaden filters) | webpart-builder | — | [x] |
| 5.4.2 | Implement empty state for no filters available | filter-builder | — | [x] |
| 5.4.3 | Implement error state for search API failure: user-friendly message + retry button | webpart-builder | — | [x] |
| 5.4.4 | Implement error state for network timeout | webpart-builder | — | [x] |
| 5.4.5 | Implement degraded state for missing permissions (e.g., no Graph access for People layout) | webpart-builder | — | [x] |
| 5.4.6 | Review and test all ErrorBoundary fallback UIs | testing | — | [x] |

---

### Step 5.5 — Performance Profiling

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.5.1 | Profile initial page load time — target < 3 seconds on warm cache | testing | §4.3 | [ ] |
| 5.5.2 | Profile search execution time — target < 1 second query-to-render | testing | §4.3 | [ ] |
| 5.5.3 | Profile DataGrid rendering with 500+ results — virtual scrolling smooth | testing | §4.3 | [ ] |
| 5.5.4 | Profile memory usage — no leaks on repeated searches, vertical switches | testing | §4.3 | [ ] |
| 5.5.5 | Optimize: identify and fix any React re-render cascades (React DevTools Profiler) | testing | §4.3 | [ ] |

---

### Step 5.6 — Documentation

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.6.1 | Admin guide: property pane configuration for each web part | — | §9 | [x] |
| 5.6.2 | Extensibility guide: how to register custom DataProvider, SuggestionProvider, ActionProvider, Layout, FilterType | — | §4.4 | [x] |
| 5.6.3 | Deployment guide: build, deploy .sppkg, run provisioning script, verify | — | §7 | [x] |
| 5.6.4 | Provisioning script documentation: prerequisites, parameters, troubleshooting | — | §7.2 | [x] |

---

### Step 5.7 — Final Validation

| # | Task | Agent | Req § | Status |
|---|------|-------|-------|--------|
| 5.7.1 | Full regression test in SPFx hosted workbench (not local) | testing | — | [ ] |
| 5.7.2 | Test on Edge, Chrome, Firefox, Safari | testing | §10.3 | [ ] |
| 5.7.3 | Deploy to test site collection, verify real SharePoint search results | testing | §7 | [ ] |
| 5.7.4 | Verify hidden list provisioning on clean site | testing | §7.2 | [ ] |
| 5.7.5 | Load test: 100+ results, 20+ filter values, 5+ verticals | testing | §4.3 | [ ] |
| 5.7.6 | Production build: `gulp bundle --ship && gulp package-solution --ship` — no errors, no warnings | testing | §7.3 | [x] |

---

## Summary

| Phase | Steps | Tasks | Focus |
|-------|-------|-------|-------|
| **Phase 1** | 14 steps | ~85 tasks | Foundation: scaffolding, store, providers, basic web parts |
| **Phase 2** | 7 steps | ~45 tasks | Rich Layouts: DataGrid, Card, People, Gallery, Detail Panel |
| **Phase 3** | 10 steps | ~40 tasks | User Features: Search Manager, sharing, collections, history |
| **Phase 4** | 9 steps | ~40 tasks | Power Features: query builder, advanced filters, bulk actions |
| **Phase 5** | 7 steps | ~30 tasks | Polish: optimization, accessibility, responsive, documentation |
| **Total** | **47 steps** | **~240 tasks** | |

---

### Full Code Audit (Post-Phase 5)

> Comprehensive audit of all code: store, providers, services, web parts, interfaces. All issues found and fixed.

| # | Issue | Severity | File(s) | Status |
|---|-------|----------|---------|--------|
| A.1 | QueryBuilder startswith/endswith generates invalid KQL (missing quotes around wildcards) | CRITICAL | `QueryBuilder.tsx` | [x] |
| A.2 | TrendingQueryProvider queries ALL SearchHistory without Author filter — list threshold issue | CRITICAL | `TrendingQueryProvider.ts` | [x] |
| A.3 | PromotedResultsService NaN date comparison — invalid dates silently pass | CRITICAL | `PromotedResultsService.ts` | [x] |
| A.4 | SchemaService + AudienceService don't check response.ok on SPContext.http.get() | CRITICAL | `SchemaService.ts`, `AudienceService.ts` | [x] |
| A.5 | SearchManagerService._currentUserId stays 0 on init failure — orphaned writes | CRITICAL | `SearchManagerService.ts` | [x] |
| A.6 | Registry.freeze() never called — registries not locked after first search | CRITICAL | `SearchOrchestrator.ts` | [x] |
| A.7 | initializeSearchContext() race condition — concurrent calls create duplicates | CRITICAL | `storeRegistry.ts` | [x] |
| A.8 | ActiveFilterPillBar infinite useEffect loop — displayMap in dependency array | HIGH | `ActiveFilterPillBar.tsx` | [x] |
| A.9 | URL sync not wired up — createUrlSyncSubscription exported but never called | HIGH | `storeRegistry.ts` | [x] |
| A.10 | FilterGroup direct imports of heavy filter components — bypassed lazy loading | HIGH | `FilterGroup.tsx` | [x] |
| A.11 | ListLayout XSS via dangerouslySetInnerHTML without sanitization | HIGH | `ListLayout.tsx` | [x] |
| A.12 | BulkActionsToolbar error missing role="alert" | MEDIUM | `BulkActionsToolbar.tsx` | [x] |
| A.13 | Missing aria-expanded on panel toggle buttons | MEDIUM | `SpSearchBox.tsx` | [x] |
| A.14 | Missing aria-live/role="status" on result count and empty states | MEDIUM | Multiple web parts | [x] |
| A.15 | Missing aria-hidden on pagination ellipsis | LOW | `Pagination.tsx` | [x] |

---

## Dependencies Between Phases

```
Phase 1.0-1.2 (Scaffold + Interfaces + Registries)
       ↓
Phase 1.3 (Store + Library Component)  ← EVERYTHING depends on this
       ↓
Phase 1.4 (URL Sync) + Phase 1.5 (Token/Search Service) + Phase 1.6 (SP Provider)
       ↓
Phase 1.7 (Orchestrator)  ← Ties store to providers
       ↓
Phase 1.8-1.11 (All 4 web parts)  ← Can be built in parallel
       ↓
Phase 1.12 (Provisioning)  ← Independent, can start anytime
       ↓
Phase 2 (Rich Layouts)  ← Depends on Phase 1 web parts being functional
       ↓
Phase 3 (User Features)  ← Depends on Phase 1 + provisioned lists
       ↓
Phase 4 (Power Features)  ← Depends on Phase 1-3 infrastructure
       ↓
Phase 5 (Polish)  ← Depends on all features being implemented
```
