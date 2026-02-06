# SP Search Project — Continuation Prompt

Use this prompt to continue development of the SP Search project with Claude Code.

---

## Project Context

**SP Search** is an enterprise SharePoint search solution built with SPFx 1.21.1. It consists of:
- **1 Library Component**: `sp-search-store` — Zustand store registry, providers, services
- **5 Web Parts**: SearchBox, SearchResults, SearchFilters, SearchVerticals, SearchManager

**Repository**: https://github.com/hmane/sp-search

**Key Files**:
- `CLAUDE.md` — Primary development guide with architecture, patterns, and rules
- `docs/sp-search-requirements.md` — Full requirements specification (1,897 lines)
- `docs/DEVELOPMENT-PLAN.md` — Step-by-step implementation plan with status tracking
- `docs/admin-guide.md` — Property pane configuration for all 5 web parts
- `docs/extensibility-guide.md` — All 5 extension points with interfaces and examples
- `docs/deployment-guide.md` — Build, deploy, provision, verify steps with bundle analysis
- `docs/provisioning-guide.md` — Script parameters, list schemas, permissions, seed data

**Local Dependencies**:
- `spfx-toolkit` at `/Users/hemantmane/Development/spfx-toolkit`

---

## What's Complete

### Phase 1: Foundation ✅
All 13 steps complete — scaffolding, interfaces, registries, Zustand store, URL sync middleware, Token/Search services, SharePoint provider, orchestrator, all 5 web parts (basic), hidden list provisioning script. Bundle tree-shaking verified (1.13.7).

### Phase 2: Rich Layouts ✅
Steps 2.1-2.6 complete — DataGrid Layout (DevExtreme lazy-loaded), Card Layout (spfx-toolkit), People Layout, Document Gallery Layout, Result Detail Panel (WOPI preview, metadata, version history, related docs), Layout Switcher, Refiner Stability mode.

### Phase 3: User Features ✅
Steps 3.1-3.9 complete — SearchManagerService (full CRUD with isReady guards, retryable init), item-level permissions (breakRoleInheritance + addRoleAssignment), Search Manager web part (standalone + panel), ShareSearchDialog (URL/Email/Teams/Users tabs), history auto-logging + click tracking + SHA-256 dedup, StateId deep link fallback (?sid=), Promoted Results / Best Bets with audience targeting, Active Filter Pill Bar (async value resolution), RecentSearchProvider, SuggestionDropdown (grouped, keyboard nav, ARIA).

### Phase 4: Power Features ✅
All 9 steps complete:
- **4.1** Advanced Filter Types — All 7 types: Checkbox, DateRange, Slider, PeoplePicker, TaxonomyTree, TagBox, Toggle + 6 formatters
- **4.2** Visual Query Builder — expandable panel with property/operator/value builders, KQL conversion
- **4.3** Visual Filter Builder — DevExtreme FilterBuilder-style AND/OR group editor
- **4.4** Bulk Actions Toolbar — 8 built-in actions: Open, Preview, Share, Pin, CopyLink, Download, Compare, ExportCsv
- **4.5** Smart Suggestions — RecentSearchProvider, TrendingQueryProvider, ManagedPropertyProvider
- **4.6** Result Annotations / Tags — personal + shared tags, TagBadges in layouts, tag-based filtering
- **4.7** Audience Targeting — Azure AD group resolution, verticals + promoted results visibility
- **4.8** Schema Helper — SchemaService (sessionStorage cache), SchemaHelperControl (Panel with Pivot tabs, DetailsList), PropertyPaneSchemaHelper factory, integrated into SpSearchResultsWebPart
- **4.9** Phase 4 integration testing verified

### Phase 5: Polish & Optimization (Partially Complete)
- **5.1** Bundle Optimization ✅ — tree-shaking verified, .sppkg at 19 MB, lazy-loaded chunks confirmed
- **5.2** Accessibility Audit ✅ — WCAG 2.1 AA: keyboard nav, screen reader, focus management, ARIA labels, live regions
- **5.4** Error Handling & Empty States ✅ — all empty/error/degraded states implemented
- **5.6** Documentation ✅ — admin guide, extensibility guide, deployment guide, provisioning script docs
- **5.7.6** Production build ✅ — `gulp bundle --ship` (0 errors, 19s) + `gulp package-solution --ship` (0 errors, 6.5s)

### Full Code Audit ✅
15 issues found and fixed (7 CRITICAL, 4 HIGH, 3 MEDIUM, 1 LOW) — see DEVELOPMENT-PLAN.md audit section.

### SPContext Refactor ✅
All services/providers use `SPContext.sp` (no direct SPFI injection), PnP side-effect imports via spfx-toolkit bundles, type-only imports for PnP types.

---

## Remaining Tasks (All Require SharePoint Deployment)

### Integration Testing — 19 tasks
```
[ ] 1.13.1-6  Phase 1: end-to-end search, filter flow, vertical switch, URL sync, multi-instance, abort
[ ] 2.7.1-6   Phase 2: 6 layouts, layout switcher, DataGrid features, Detail Panel, refiner stability, lazy loading
[ ] 3.10.1-7  Phase 3: save/load/delete, sharing, collections, history, promoted results, StateId, pill bar
```

### Responsive Design (5.3) — 5 tasks
```
[ ] 5.3.1 Test all layouts at mobile (320px), tablet (768px), desktop (1024px+)
[ ] 5.3.2 Search Box: full width on mobile, inline scope selector collapse
[ ] 5.3.3 Filters: collapse to panel/drawer on mobile
[ ] 5.3.4 Verticals: overflow to "More" dropdown on narrow screens
[ ] 5.3.5 Detail Panel: full-screen on mobile
```

### Performance Profiling (5.5) — 5 tasks
```
[ ] 5.5.1 Profile initial page load — target < 3 seconds (warm cache)
[ ] 5.5.2 Profile search execution — target < 1 second query-to-render
[ ] 5.5.3 Profile DataGrid with 500+ results — virtual scrolling smooth
[ ] 5.5.4 Profile memory usage — no leaks on repeated searches
[ ] 5.5.5 Optimize React re-render cascades (React DevTools Profiler)
```

### Final Validation (5.7) — 5 tasks
```
[ ] 5.7.1 Full regression in SPFx hosted workbench
[ ] 5.7.2 Cross-browser: Edge, Chrome, Firefox, Safari
[ ] 5.7.3 Deploy to test site, verify real SharePoint search
[ ] 5.7.4 Verify hidden list provisioning on clean site
[ ] 5.7.5 Load test: 100+ results, 20+ filter values, 5+ verticals
```

**Total remaining: 34 tasks — all require SharePoint Online environment.**

---

## Bundle Size Reference (Ship Build)

| Entry Bundle | Size |
|---|---|
| sp-search-store-library | 14 KB |
| sp-search-verticals-web-part | 775 KB |
| sp-search-box-web-part | 1.0 MB |
| sp-search-results-web-part | 1.1 MB |
| sp-search-filters-web-part | 1.4 MB |
| sp-search-manager-web-part | 1.4 MB |

Lazy-loaded chunks: CardLayout (122 KB), DevExtremeDataGrid (71 KB), TaxonomyTree (50 KB), SearchManager (46 KB), VisualFilterBuilder (45 KB), DetailPanel (27 KB), PeopleLayout (25 KB), PeoplePickerFilter (17 KB), TagBoxFilter (12 KB), GalleryLayout (3 KB), DataGridLayout wrapper (2.1 KB).

.sppkg total: 19 MB.

---

## Critical Implementation Rules

1. **Import Patterns** (MANDATORY):
   ```typescript
   // spfx-toolkit — ALWAYS direct path imports
   import { Card } from 'spfx-toolkit/lib/components/Card';
   // NEVER: import { Card } from 'spfx-toolkit';

   // spfx-toolkit — PnP side-effect imports via bundles
   import 'spfx-toolkit/lib/utilities/context/pnpImports/lists';
   import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
   import 'spfx-toolkit/lib/utilities/context/pnpImports/security';
   // NEVER: import '@pnp/sp/lists' (use toolkit bundles instead)

   // PnP types — ALWAYS type-only imports
   import type { ISearchQuery } from '@pnp/sp/search';
   // NEVER: import { ISearchQuery } from '@pnp/sp/search';

   // SPContext — use SPContext.sp instead of injecting SPFI
   import { SPContext } from 'spfx-toolkit/lib/utilities/context';
   // Initialize in onInit: await SPContext.basic(this.context, 'Name');
   // Use anywhere: SPContext.sp.web.lists...

   // Fluent UI v8 — ALWAYS tree-shakable imports
   import { Panel } from '@fluentui/react/lib/Panel';
   // NEVER: import { Panel } from '@fluentui/react';

   // DevExtreme — Lazy load heavy components
   const DataGrid = React.lazy(() => import('devextreme-react/data-grid'));

   // createLazyComponent — cast as any due to @types/react mismatch
   const Foo: any = createLazyComponent(() => import('./Foo') as any, { errorMessage: '...' });
   // Do NOT wrap output in <React.Suspense> — it bundles Suspense internally
   ```

2. **Architecture Rules**:
   - Web parts NEVER call PnPjs/Graph directly — always through ISearchDataProvider
   - All inter-webpart communication via Zustand store — no SPFx Dynamic Data
   - AbortController on every search — cancel in-flight requests before new ones
   - Registries freeze after first search execution (SearchOrchestrator enforces)
   - Store init uses promise-based lock to prevent race conditions
   - URL sync wired in storeRegistry via `createUrlSyncSubscription`
   - Services use `SPContext.sp` (no SPFI constructor injection)
   - Use `createSPExtractor` for type-safe field extraction from list items
   - Use `BatchBuilder` for batch SharePoint operations

3. **Data Rules**:
   - SearchHistory list WILL exceed 5,000 items — ALWAYS filter by Author FIRST in CAML
   - CollapseSpecification fails SILENTLY on non-sortable properties — validate before sending
   - Taxonomy refiners use GP0|#GUID format — must resolve to labels via PnP Taxonomy API
   - Date refiners MUST use FQL range() — NOT raw KQL date comparisons
   - SearchConfiguration uses `ConfigValue` column (NOT `ConfigData`)
   - Promoted results: one rule per list item with matchType/matchValue/promotedItems structure
   - SearchManagerService has isReady guards on ALL write methods, retryable initialize()

4. **Path Aliases** (configured in tsconfig.json):
   ```
   @store/*        → src/libraries/spSearchStore/*
   @interfaces/*   → src/libraries/spSearchStore/interfaces/*
   @services/*     → src/libraries/spSearchStore/services/*
   @providers/*    → src/libraries/spSearchStore/providers/*
   @registries/*   → src/libraries/spSearchStore/registries/*
   @orchestrator/* → src/libraries/spSearchStore/orchestrator/*
   @webparts/*     → src/webparts/*
   ```

---

## Commands

```bash
# Development
gulp serve                                           # Start local workbench
gulp bundle --ship && gulp package-solution --ship   # Production build

# Testing
npx jest                                             # Run unit tests
npx jest --watch                                     # Watch mode

# spfx-toolkit (rebuild if needed)
cd /Users/hemantmane/Development/spfx-toolkit && npm run build
```

---

## Starting Point

To continue development:

1. Read `CLAUDE.md` for full architecture context
2. Check `docs/DEVELOPMENT-PLAN.md` for current status
3. All remaining tasks require SharePoint Online deployment:
   - Deploy .sppkg to App Catalog
   - Run `scripts/Provision-SPSearchLists.ps1` on target site
   - Integration testing in SPFx hosted workbench
   - Responsive testing at various breakpoints
   - Performance profiling with Chrome DevTools
   - Cross-browser validation

**All local development is complete.** Next step is deployment to a test SharePoint site for integration testing and validation.
