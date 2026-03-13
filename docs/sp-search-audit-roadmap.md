# SP Search Audit And Enhancement Roadmap

## Executive Summary

SP Search has a strong foundation. The shared Zustand store, multi-web-part search context, URL sync, extensibility registries, and Search Manager concepts are materially better than a basic PnP Modern Search replacement. The solution already has the shape of a reusable enterprise search platform.

It is not yet a reliable "one stop" search experience for any SharePoint site. The main issue is not architecture. The main issue is product completion and operational hardening. Several advanced features are partially implemented, some advertised capabilities are not fully wired through the runtime, and the current loading behavior creates avoidable trust problems for end users.

From a power SharePoint search user perspective, the highest-value path is:

- Fix perceived reliability first: loading states, grid failures, broken test harness, mobile and slow-network behavior.
- Simplify the default experience: lead with List, Compact, and Grid; treat People and Gallery as opt-in vertical-specific layouts.
- Finish the incomplete platform promises before adding more surface area.
- Add admin presets so a site owner can stand up a good search page without understanding KQL, managed properties, and result sources in depth.

## Scope Of This Audit

This document is based on a code audit of the current SPFx solution and a local validation pass.

Observed locally:

- `npm run type-check`: passed
- `npm run build`: passed
- `npm test -- --runInBand`: failed before executing tests because `ts-jest` cannot resolve `jest-util`

Key files reviewed:

- `src/webparts/spSearchResults/components/SpSearchResults.tsx`
- `src/webparts/spSearchResults/components/DataGridLayout.tsx`
- `src/webparts/spSearchResults/components/DataGridContent.tsx`
- `src/webparts/spSearchResults/SpSearchResultsWebPart.ts`
- `src/webparts/spSearchBox/components/SpSearchBox.tsx`
- `src/webparts/spSearchFilters/SpSearchFiltersWebPart.ts`
- `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts`
- `src/libraries/spSearchStore/providers/SharePointSearchProvider.ts`
- `src/libraries/spSearchStore/store/createStore.ts`
- `src/libraries/spSearchStore/store/slices/resultSlice.ts`

## What Is Already Strong

These are the parts worth keeping and strengthening, not redesigning:

- Shared search context via library component and Zustand store. This is cleaner than SPFx Dynamic Data for coordinated search experiences.
- Multi-instance isolation via `searchContextId`. This is the right pattern for pages with more than one search experience.
- URL synchronization and state restore. This is essential for deep-linkable enterprise search.
- Search Manager concept: saved searches, history, collections, and sharing are meaningful differentiators.
- Result detail panel and active filter pill bar. These are useful daily features, not gimmicks.
- Filter type registry and schema helper. These make the platform more maintainable for administrators.
- Vertical model with audience targeting and count support. This is the right direction for site-specific and role-specific search pages.

## Priority Findings

### 1. First-load UX is weak and can show the wrong state

Impact:

- Users see "No results found" during initial load or between state transitions.
- This undermines trust immediately, especially on slower pages and cold caches.

Evidence:

- `SpSearchResults.tsx` renders `EmptyState` whenever `items.length === 0` and `isLoading` is false.
- `createStore.ts` resets result state with `isLoading: false`.
- `SpSearchResults.tsx` also contains a local `emptyState` fallback with `isLoading: false`.
- `SpSearchResults.tsx` initializes its local subscribed snapshot with `isLoading: false`, creating a small but real race window before the first store subscription update.

Recommendation:

- Introduce distinct UI states: `idle`, `loading-initial`, `loading-refresh`, `loaded-empty`, `loaded`.
- Add a `hasExecutedSearch` or `hasHydratedResults` flag in the Zustand store, ideally in `resultSlice`, instead of inferring state from `items.length`.
- On first load, show a skeleton or neutral "Preparing search..." state, never "No results found."
- On subsequent searches, keep previous results visible and show a lightweight loading overlay instead of clearing the canvas. The store flag fixes first-load trust; retained-results behavior fixes refresh-search trust.

Expected outcome:

- The page feels fast even when the network is not.
- Users stop seeing contradictory UI during load.

### 2. Data grid failure handling is too generic, and likely masks the real runtime error

Impact:

- Users see "Failed to load data grid layout" with no actionable detail.
- Support and debugging are harder than they need to be.

Evidence:

- `SpSearchResults.tsx` lazy-loads `DataGridLayout` with `createLazyComponent(..., { errorMessage: 'Failed to load data grid layout' })`.
- `DataGridLayout.tsx` then lazy-loads `DataGridContent` again with `React.lazy`.
- `DataGridLayout.tsx` also wraps the inner lazy component with `React.Suspense` directly, even though `createLazyComponent` already provides Suspense and error boundary behavior.
- The outer lazy-load error boundary catches descendant render errors too, so a DevExtreme render failure can appear as a layout chunk failure.
- DevExtreme CSS is loaded from `cdn3.devexpress.com` at runtime in both `SpSearchResultsWebPart.ts` and `SpSearchFiltersWebPart.ts`.

Inference:

- The user-facing message is probably not identifying the true fault. `React.Suspense` handles pending lazy loads, but it does not catch render errors. A nested runtime error inside `DataGridContent` can therefore bubble to the outer `createLazyComponent` error boundary and still surface as "Failed to load data grid layout."

Recommendation:

- Remove the double lazy-loading chain. Lazy-load the grid once.
- Add a grid-specific error boundary that logs the real error message and stack.
- If grid rendering fails, fall back to List layout automatically and show a non-blocking message.
- Stop depending on external CDN CSS for DevExtreme in production. In SPFx terms, either package the CSS via the solution asset pipeline or copy it to the same tenant-controlled Azure Storage or SharePoint CDN location used for solution assets, instead of loading it from `cdn3.devexpress.com` at runtime.
- Preload the grid chunk when the user hovers the layout switcher or when Grid is the default layout.

Expected outcome:

- Grid problems become diagnosable.
- Users still get results even when the advanced view fails.

### 3. The implemented DataGrid is materially below the advertised feature set

Impact:

- The current grid is not yet a power-user replacement for a true search results table.
- The gap between expected and actual behavior will create dissatisfaction.

Evidence:

- `DataGridContent.tsx` currently renders a simple grid with six columns and row click behavior.
- It does not implement the spec-level capabilities described elsewhere in the repo such as virtual scrolling, grouping, column chooser, export, filtering row, selection, fixed columns, or responsive grid behavior.

Recommendation:

- Reposition the current grid as a "basic table view" until the advanced features are real.
- For an enterprise-ready Grid milestone, add:
  - column configuration from selected properties
  - filter row
  - column chooser
  - virtual scrolling
  - stable row selection
  - CSV/Excel export
  - persisted user column preferences

Expected outcome:

- The grid becomes worth using for compliance, records, finance, policy, and reporting scenarios.

### 4. Several configured or documented capabilities are not fully wired into the runtime

Impact:

- Admins will configure options that do nothing.
- The platform appears broader than it is.

Evidence and examples:

- `queryInputTransformation` is passed into `SpSearchBox` props but is not used in `SpSearchBox.tsx`.
- `queryInputTransformation` belongs in the search execution path, not the input component. It should be applied in `SearchOrchestrator.ts` as part of query construction before the final KQL query is built, replacing tokens such as `{searchTerms}` with the current user input.
- `operatorBetweenFilters` exists in the Filters web part model and property pane, but is not used in filtering logic.
- `IVerticalDefinition.dataProviderId` exists, but `SearchOrchestrator._getProvider()` always returns the first provider and does not honor per-vertical provider selection.
- `bulkSelection` and bulk action plumbing exist in the store, but are not surfaced in the results experience.
- `bulkSelection` lives in `uiSlice`, so selection can persist across layout switches even when the active layout does not render visible selection UI. That creates a predictable hidden-selection bug when moving between Grid and non-selectable layouts.
- The requirements and docs describe configurable layout availability, but the runtime toolbar always shows all six layouts and the property pane only supports `defaultLayout`.

Recommendation:

- Treat these as a platform integrity backlog and close them before adding new features.
- Mark each item as one of:
  - implement now
  - hide from property panes and docs
  - defer explicitly with clear release notes

Expected outcome:

- Site owners stop encountering "checkboxes that do nothing."
- The product becomes easier to trust and easier to support.

### 5. The test harness is broken, and the repo currently has no working automated safety net

Impact:

- Refactoring will be risky.
- Search reliability bugs will keep recurring.

Evidence:

- `jest.config.js` expects tests under `tests`.
- `npm test -- --runInBand` fails because `ts-jest` cannot resolve `jest-util`.
- No tests were actually executed during the local run.

Recommendation:

- Fix the Jest dependency chain first.
- Add a small but meaningful test suite:
  - orchestrator request lifecycle and abort behavior
  - URL sync encode/decode
  - refiner merging stability
  - first-load and zero-result state transitions
  - provider selection by vertical

Expected outcome:

- The next round of fixes can be made with confidence.

## Product Recommendations From A Power SharePoint Search User Perspective

### Keep As Core Features

- List layout as the default search result experience
- Compact layout for dense scanning
- DataGrid as an advanced opt-in power-user view
- Filters with active pill bar
- Verticals with meaningful business scopes
- Result detail panel with preview and metadata
- Saved searches and recent history

These are features people will actually use repeatedly.

### Deprioritize By Default

- People layout
- Gallery layout
- Standalone Search Manager web part on normal search pages
- Visual filter builder for general users

Reasoning:

- People search only becomes truly valuable when backed by Microsoft Graph data, better profile fields, and presence-aware actions. In the current implementation it is mostly a SharePoint-result card view, not a differentiated people experience.
- Gallery is useful for image-heavy or asset-heavy repositories, but it should not be a default layout for general SharePoint search.
- Search Manager is valuable, but as a panel launched from the Search Box it is usually enough. A dedicated web part makes sense only on a specialist search hub or admin page.
- Visual query/filter builders are useful for a small audience. They should not dominate the default UX.

### Recommended Default Product Shape

For a general-purpose SharePoint site search page, I would standardize on:

- Search Box with suggestions, scope, recent searches, and Search Manager panel access
- Verticals: All, Documents, Pages, Sites, People
- Result layouts enabled by default: List, Compact, Grid
- Filters: file type, site/location, author/owner, modified date, content type, department if available
- Detail panel enabled by default
- Saved searches enabled

I would only enable:

- People layout on a dedicated people vertical after Graph provider support is complete
- Gallery layout on image/video/brand-asset search pages

## What Is Needed To Make This A True "One Stop Search" Solution

### 1. Opinionated page templates and presets

Today the platform is powerful but admin-heavy. To scale across many sites, it needs presets such as:

- Site Search
- Hub Search
- Policy Search
- Knowledge Base Search
- People Search
- Document Center Search

Each preset should configure:

- default verticals
- default filters
- default selected properties
- default sort options
- default layout set
- scope behavior

### 2. Better admin guardrails

Add admin-time validation for:

- invalid managed properties
- missing selected properties required by the active layout
- result source GUID format
- filters configured for non-refinable properties
- layout/vertical combinations that do not make sense

Also add a health-check page or panel that reports:

- provider registered
- search context initialized
- URL sync active
- current scope and query template
- missing property mappings
- last search duration and error

### 3. Better zero-result recovery

When a search returns no results, the product should help recover:

- show suggested alternate queries
- show recently successful searches
- offer "clear filters" prominently
- show scope and filter summary
- optionally show popular content or curated links

### 4. Real people search support

If People is going to remain a first-class experience, complete the Graph-backed path:

- Graph provider for people vertical
- presence and richer profile data
- manager/team/org context
- mail/chat/profile actions

Without this, People view should be treated as secondary.

### 5. Relevance and analytics loop

To become the default search experience across sites, the solution needs feedback loops:

- zero-result analytics
- top queries
- abandoned queries
- click-through by vertical
- promoted result effectiveness
- filter usage analytics

These metrics should drive tuning, not just exist as data.

## Recommended Roadmap

### Sprint 0: Stabilize The Current Product

Deliver in the next sprint:

1. Fix first-load empty-state flash.
2. Fix or properly isolate DataGrid load failures.
3. Bundle or tenant-host DevExtreme CSS instead of relying on external CDN.
4. Repair Jest and add smoke tests for orchestrator and URL sync.
5. Add runtime logging around provider selection, lazy-load failures, and first search timing.

Dependency guidance:

- Item 4 should start first or in parallel, because it restores the safety net for the other stabilization changes.
- Items 1, 2, and 3 are largely independent workstreams and can be handled by different contributors in parallel.
- Item 1 is mostly store and results-component work.
- Item 2 is mostly lazy-loading, error-boundary, and DataGrid integration work.
- Item 3 is packaging and asset-delivery work specific to the SPFx deployment path.
- Item 5 fits best after Items 1 and 2 begin, so logging can be added around the exact states and failure paths being fixed.

### Sprint 1: Clean Up Feature Drift

1. Implement or remove `queryInputTransformation`.
2. Implement or remove `operatorBetweenFilters`.
3. Implement per-vertical `dataProviderId` support.
4. Implement configurable layout availability and hide irrelevant layouts by vertical.
5. Either surface bulk actions fully or remove the unfinished plumbing from messaging.

### Sprint 2: Ship A Better Default Search Experience

1. Introduce layout presets by scenario.
2. Narrow the default layout switcher to List, Compact, Grid.
3. Improve slow-loading UX with retained results plus loading overlay.
4. Add a meaningful zero-results recovery panel.
5. Add admin validation and setup diagnostics.

### Sprint 3: Become The Preferred Enterprise Search Layer

1. Finish advanced DataGrid capabilities.
2. Complete Graph-backed People vertical.
3. Add analytics and tuning dashboard.
4. Add preset-based provisioning scripts for common site archetypes.
5. Add performance hardening for 500+ result scenarios and mobile responsiveness.

---

## Sprint 3 Detailed Scope

Sprint 2 closed the baseline gaps: trustworthy loading states, zero-results recovery, opinionated layout defaults, admin validation, and layout/URL consistency. Sprint 3 completes the three platform promises that differentiate SP Search from PnP Modern Search: a power-user data grid, a Graph-backed people experience, and an analytics feedback loop. Provisioning presets and mobile hardening follow when those tracks are stable.

### Sequencing Overview

Run Tracks A and B in parallel. Track C begins once Track A is reviewable. Tracks D and E run across the sprint as fill-in work and polish.

```
Week 1-2:   Track A  (DataGrid completion)
            Track B  (Graph People — provider + layout)
Week 2-3:   Track C  (Analytics light: zero-result logging, health panel)
            Track E  (Mobile hardening — can start any time)
Week 3-4:   Track D  (Provisioning presets — needs A+B stable)
            Track C  (Full analytics: top queries, click-through)
```

---

### Track A — Advanced DataGrid (Item 1)

**Goal:** Turn the current "basic table view" into the power-user grid promised in the spec, without rebuilding the DevExtreme foundation.

**Deliverables — in order:**

**A1. DataGrid-specific error boundary with List fallback**
- Add a dedicated React error boundary inside `DataGridLayout.tsx` (not relying on the outer `createLazyComponent` boundary).
- On error: automatically switch `activeLayoutKey` to `'list'` via the store, then show a non-blocking `MessageBar` informing the user the grid failed and they've been moved to List view.
- Log the real error message and stack to `console.error` so it is visible in browser DevTools.
- Remove the double lazy-load chain (`DataGridLayout` → `DataGridContent`): the outer `createLazyComponent` is sufficient; `DataGridContent` should be imported directly inside `DataGridLayout`.

**A2. Column configuration from `selectedPropertiesCollection`**
- `DataGridContent` currently hardcodes six columns. Replace with columns derived from `store.getState().selectedProperties` (the same comma-separated list the admin configured in the property pane).
- Map each property name to a typed `DataGrid` column definition using the `cellRenderers` registry already present in the codebase.
- Fall back to auto-generated text columns for properties without a registered renderer.

**A3. Filter row, column chooser, and virtual scrolling**
- Enable DevExtreme's built-in `filterRow`, `columnChooser`, and `scrolling.mode='virtual'` via props.
- Virtual scrolling replaces per-page navigation for the DataGrid view only; the main paginated store state is unaffected.
- Guard: virtual scrolling is only activated when `pageSize >= 25` and `totalCount > pageSize` (i.e., when there are multiple pages of data).

**A4. Stable row selection + bulk action surface**
- Wire `DataGrid` `selection.mode='multiple'` to the store's `bulkSelection` slice.
- When the active layout switches away from Grid, clear `bulkSelection` to prevent the hidden-selection bug identified in the audit (`bulkSelection` persisting across layout switches).
- Surface a `BulkActionsToolbar` above the grid when `bulkSelection.length > 0`. Minimum actions: Copy link, Download (if `FileRef` is present), Clear selection.
- Do not add more bulk actions in this sprint; the toolbar surface is the deliverable, not a full action set.

**A5. CSV export**
- Add an Export button to the DataGrid toolbar (visible only in Grid layout).
- Export the current page's results to CSV using the column configuration from A2.
- For virtual scrolling mode, export what is currently loaded, not the full result set.
- No server-side export in this sprint.

**A6. Persisted user column preferences**
- Save `columnChooser` visibility and column order to `localStorage` keyed by `searchContextId + '-grid-columns'`.
- Restore on mount. Invalid or stale preferences should be silently discarded (no error boundary needed here).

**Out of scope for Sprint 3:** Excel (XLSX) export, server-side export of all pages, row grouping, fixed columns, print view.

---

### Track B — Graph-backed People Vertical (Item 2)

**Goal:** Make the People layout actually differentiated, not just a card view over SharePoint results.

**Deliverables — in order:**

**B1. `GraphSearchProvider` implementation**
- Implement `GraphSearchProvider` in `src/libraries/spSearchStore/providers/data/GraphSearchProvider.ts`.
- Target entity type: `person` via Microsoft Search (`/search/query` with `entityTypes: ['person']`).
- Map Graph person result fields to `ISearchResult`: `displayName`, `mail`, `jobTitle`, `department`, `officeLocation`, `userPrincipalName`, `id`.
- Respect `pageSize` and `currentPage` from the store. Graph Search supports `from`/`size` pagination.
- Propagate `abortController` so in-flight requests are cancelled on new queries (same contract as `SharePointSearchProvider`).
- Presence: batch presence lookups via `/communications/presences` for the current page's user IDs. Cache results in a `Map<string, PresenceStatus>` local to the provider instance.

**B2. Per-vertical `dataProviderId` wiring**
- Sprint 1 already exposed `IVerticalDefinition.dataProviderId`. Verify `SearchOrchestrator._getProvider()` actually reads it. If not (current evidence suggests it always returns the first provider), implement the lookup: `dataProviders.get(vertical.dataProviderId) ?? dataProviders.getFirst()`.
- This is required before `GraphSearchProvider` can be selected for a People vertical.

**B3. People layout enhancements**
- Add presence badge to `UserPersona` in `PeopleLayout.tsx` using the presence data surfaced by `GraphSearchProvider`.
- Add action buttons per person card: Email (`mailto:`), Teams chat (`https://teams.microsoft.com/l/chat/0/0?users=...`), View profile (SharePoint profile URL).
- Show department and office location below job title when populated.
- People layout should degrade gracefully when backed by `SharePointSearchProvider` instead of `GraphSearchProvider`: hide presence badge and Teams action if those fields are missing.

**B4. `GraphSearchProvider` registration in the Verticals web part**
- In `SpSearchVerticalsWebPart.ts` (or wherever verticals are configured), allow the admin to set `dataProviderId: 'graph'` per vertical via the property pane.
- Document in the property pane description: "Use 'graph' for People verticals. Requires Microsoft Graph search permissions (Sites.Read.All is not sufficient; People.Read is required)."

**Out of scope for Sprint 3:** org chart view, manager/team traversal, advanced profile enrichment from Delve/Viva, audience-targeted results via Graph.

---

### Track C — Analytics and Tuning (Item 3)

**Goal:** Give admins visibility into search quality. Start with zero-operational-overhead instrumentation, not a full BI pipeline.

**Deliverables — in order:**

**C1. Zero-result query logging**
- When a search completes with `totalCount === 0`, write a record to the `SearchHistory` SharePoint list (already provisioned) with a `ZeroResult` flag field.
- The list write is fire-and-forget — do not block the UI or retry on failure.
- Cap writes: if the same query text logged a zero-result in the same session (in-memory `Set`), skip the duplicate write.
- **Data rule:** Always include `Author eq [Me]` as the first CAML predicate on any list query against `SearchHistory` (5,000-item threshold rule from the project memory).

**C2. Admin health-check panel in edit mode**
- Extend the existing edit-mode `MessageBar` in `SpSearchResults.tsx` into a collapsible diagnostics panel (Fluent UI `Panel` or inline expandable section).
- Show when `isEditMode === true`. Report:
  - Search context ID and whether the store is initialized.
  - Active data provider ID and whether it is registered.
  - URL sync status (active/inactive).
  - Current scope, query template, and result source ID.
  - Last search duration (ms), result count, and error message if the last search failed.
  - Whether `availableLayouts` matches `activeLayoutKey` (flags misconfiguration).
- All values come from `store.getState()` — no additional API calls.

**C3. Click-through tracking**
- When a user opens a result in the detail panel or navigates to a result URL, write a `ClickThrough` event record to `SearchHistory`.
- Fields: query text at time of click, clicked result title, clicked result URL, vertical, layout active at click time.
- Same fire-and-forget and dedup rules as C1.

**C4. Top queries and zero-result report (admin-only panel)**
- Add a "Search Insights" tab to the Search Manager panel (already exists for saved searches and history).
- Query `SearchHistory` grouped by query text, filtered to `ZeroResult eq true`, ordered by count desc, limit 20.
- Show two tables: Top 20 queries overall (by frequency) and Top 20 zero-result queries.
- Date filter: Last 7 days / Last 30 days / All time. Use CAML `<DateRangesOverlap>` or `<Geq>` on `Created`.
- No charting library — plain Fluent UI `DetailsList` is sufficient.

**Out of scope for Sprint 3:** promoted result CTR analytics, Viva/Clarity integration, external BI export, Power BI embedding.

---

### Track D — Preset-based Provisioning Scripts (Item 4)

**Goal:** A site owner can deploy a complete, working search experience for a common archetype in one script run.

**Deliverables:**

**D1. Site Search preset**
- Scope: current site collection.
- Verticals: All, Documents, Pages, People.
- Default filters: FileType, LastModifiedTime, Author.
- Default layout: List. Layout switcher: List, Compact, Grid.
- Selected properties: Title, FileRef, LastModifiedTime, Author, FileType, HitHighlightedSummary.

**D2. Document Center preset**
- Scope: current site collection.
- Verticals: All Documents, Word, Excel, PowerPoint, PDF.
- Default filters: FileType, LastModifiedTime, Author, Department (if mapped).
- Default layout: Grid (DataGrid). Layout switcher: List, Compact, Grid.
- CollapseSpecification: `FileLeafRef:1` (deduplicate co-authored files).

**D3. People Search preset**
- Scope: All SharePoint (tenant-wide).
- Verticals: People (Graph provider), Org Chart (deferred — placeholder vertical only).
- Default filters: Department, Office, JobTitle.
- Default layout: People. Layout switcher: People, List.
- `dataProviderId: 'graph'` on People vertical.

**D4. News preset**
- Scope: hub site or current collection depending on param.
- Verticals: All News, Department News.
- Result source: Pages library filter (`ContentClass:STS_ListItem_850`).
- Default layout: Card. Layout switcher: Card, List.

Each preset is a PowerShell function in `scripts/Setup-SPSearchSite.ps1` accepting `-Preset <name>` and `-SiteUrl`. The function creates the page, adds web parts, and sets all property values via the existing Add-PnPPageWebPart pattern.

**Out of scope:** Knowledge Base, Hub Search, and Policy Search presets — defer to Sprint 4.

---

### Track E — Performance Hardening and Mobile Responsiveness (Item 5)

Track E items are parallelizable filler work — pick them up when a contributor is unblocked from Tracks A–D.

**E1. Mobile layout audit**
- Test List and Compact layouts at 320px, 375px, 480px, and 768px viewport widths.
- Fix any overflow, truncation, or touch-target issues found.
- Active filter pill bar should wrap gracefully at narrow widths.
- Layout switcher toolbar should collapse to an icon-only strip below 480px.

**E2. Slow-network guard for filter loading**
- Refiners currently load with the main search. On slow networks, the filter panel may show stale options for the previous query for several seconds.
- Add a "Updating filters..." subtle loading state in `SpSearchFilters` when `isLoading === true` and `activeFilters` has changed since the last refiner response.
- Do not block the filter UI; show the previous options with a reduced-opacity overlay (same pattern as the results loading overlay from Sprint 2).

**E3. Pagination performance for 500+ result sets**
- Profile the `resultSlice` update path when `totalCount > 500`. Ensure `items` array replacement does not trigger unnecessary re-renders in the filter panel or verticals.
- If profiling reveals expensive re-renders, wrap `SpSearchFilters` and `SpSearchVerticals` in `React.memo` with explicit prop comparison.
- This is a profile-first item: do the measurement before writing any optimization code.

**E4. Preload DataGrid chunk on hover**
- When the user hovers the Grid layout button in the toolbar, call `import('./DataGridLayout')` to begin chunk preloading.
- Also preload automatically if `defaultLayout === 'grid'` during `onInit()`, so the first render of the grid does not incur a cold-load delay.

---

### Sprint 3 Exit Criteria

Sprint 3 is done when:

- [x] DataGrid columns reflect `selectedPropertiesCollection`, filter row is enabled, and CSV export works end-to-end.
- [x] A DataGrid render failure degrades to List layout with a visible (non-blocking) message, not a generic error boundary.
- [x] `GraphSearchProvider` returns people results for at least one Graph-backed vertical and presence status is shown on People cards.
- [x] Per-vertical `dataProviderId` selection routes queries to the correct provider in a live test.
- [x] Zero-result queries are logged to `SearchHistory` and appear in the Search Insights tab.
- [x] Health and Insights tabs in Search Manager surface zero-result and click-through analytics.
- [x] At least five provisioning presets (general, documents, people, news, media) available via `Search-ScenarioPresets.ps1`.
- [x] List and Compact layouts are usable without horizontal scroll on a 375px viewport.

**Sprint 3 is complete.** See `docs/sp-search-sprint3-release.md` for the full delivery
record, regression checklist, admin documentation update requirements, and deferred items.

## Keep, Fix, Deprioritize

### Keep And Strengthen

- shared store and orchestrator architecture
- saved searches, history, collections
- result detail panel
- verticals
- active filter pill bar
- schema helper and typed registries

### Fix Before Wider Rollout

- first-load empty state
- grid load failure handling
- broken test setup
- incomplete DataGrid
- dead or misleading property-pane options
- missing per-vertical provider support
- default layout set not configurable

### Deprioritize Or Hide By Default

- People layout until Graph support is real
- Gallery layout except for media-heavy scenarios
- standalone Search Manager page parts on standard site pages
- visual builder experiences for broad audiences

## Suggested Success Criteria

The solution is ready to position as a full enterprise replacement when the following are true:

- no false empty state on first load
- grid failures degrade gracefully to a safe layout
- automated tests run in CI and cover core search flows
- the default page experience is usable without deep admin tuning
- advanced features shown in the UI are actually complete
- a site owner can choose a preset and get a good search page in minutes

## Final Recommendation

Do not treat the next iteration as a feature expansion exercise. Treat it as a product-hardening release with selective UX simplification.

The fastest path to making this the preferred SharePoint search layer is:

- make the default experience simpler
- make the runtime states trustworthy
- finish the incomplete promises already exposed in code and docs
- reserve niche layouts and advanced builders for the scenarios that truly need them

If that is done well, SP Search can be more than a PnP Modern Search replacement. It can become a reusable search platform for site, hub, department, and enterprise-wide search pages.
