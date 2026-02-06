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

**Local Dependencies**:
- `spfx-toolkit` at `/Users/hemantmane/Development/spfx-toolkit`

---

## What's Complete

### Phase 1: Foundation ✅
All 13 steps complete — scaffolding, interfaces, registries, Zustand store, URL sync middleware, Token/Search services, SharePoint provider, orchestrator, all 5 web parts (basic), hidden list provisioning script.

### Phase 2: Rich Layouts ✅
Steps 2.1-2.6 complete — DataGrid Layout (DevExtreme), Card Layout (spfx-toolkit), People Layout, Document Gallery Layout, Result Detail Panel, Layout Switcher, Refiner Stability mode.

### Phase 3: User Features ✅
Steps 3.1-3.9 complete — SearchManagerService (CRUD), item-level permissions, Search Manager web part, ShareSearchDialog (4 tabs), history auto-logging + click tracking + cleanup TTL, StateId deep link fallback, Promoted Results / Best Bets, Active Filter Pill Bar, RecentSearchProvider.

### SPContext Refactor ✅
- All services/providers refactored to use `SPContext.sp` (no direct SPFI injection)
- All PnP side-effect imports replaced with `spfx-toolkit/lib/utilities/context/pnpImports/*` bundles
- PnP type imports use `import type` (zero bundle cost)
- Web parts call `SPContext.basic(this.context, 'Name')` in `onInit()`
- `SearchManagerService` uses `createSPExtractor`, `BatchBuilder`, `SPContext.logger`

### Step 4.1: Advanced Filter Types ✅
All 7 filter types implemented and registered:
- `TaxonomyTreeFilter.tsx` — DevExtreme TreeView, hierarchical taxonomy with GP0|#GUID resolution
- `PeoplePickerFilter.tsx` — PnP PeoplePicker, claim string resolution
- `SliderFilter.tsx` — DevExtreme RangeSlider, FQL range(decimal()) tokens, file size formatting
- `TagBoxFilter.tsx` — DevExtreme TagBox, tag-style multi-select
- `ToggleFilter.tsx` — Fluent UI Toggle, three-state (All/Yes/No)
- 6 formatters: Taxonomy, People, Numeric, Boolean, Date, Default
- `registerBuiltInFilterTypes.ts` — registers all 7 in FilterTypeRegistry
- `FilterGroup.tsx` updated to route all 7 filter types
- `SpSearchFiltersWebPart.ts` calls registration in `onInit()`

---

## Remaining Tasks

### Step 2.7 — Phase 2 Integration Testing (6 tasks)
```
[ ] 2.7.1 Test all 6 layouts render correctly with search results
[ ] 2.7.2 Test layout switcher: toggle between all layouts preserves results, selection, scroll position
[ ] 2.7.3 Test DataGrid: sort, filter, group, export, virtual scrolling, responsive card mode
[ ] 2.7.4 Test Result Detail Panel: opens on click, preview loads, metadata formatted, version history
[ ] 2.7.5 Test refiner stability: rapid typing doesn't cause filter options to flicker
[ ] 2.7.6 Verify DataGrid lazy loading — bundle chunk only loaded when grid layout selected
```

### Step 3.9.2 — Suggestion Dropdown (1 remaining task)
```
[ ] 3.9.2 Implement SuggestionDropdown.tsx in Search Box — dropdown showing grouped suggestions
```

### Step 3.10 — Phase 3 Integration Testing (7 tasks)
```
[ ] 3.10.1 Test save/load/delete search lifecycle
[ ] 3.10.2 Test sharing: URL, email, Teams, user-specific with item-level permissions
[ ] 3.10.3 Test collections: create, pin from result, view, share, manage
[ ] 3.10.4 Test history: auto-log, deduplication, cleanup, re-execute
[ ] 3.10.5 Test promoted results: rule matching, display, dismissal
[ ] 3.10.6 Test StateId fallback: long URL → ?sid= → state restoration
[ ] 3.10.7 Test pill bar: display, dismiss, clear all, human-readable formatting
```

### Phase 4 — Power Features (8 remaining steps, ~30 tasks)
- Step 4.2: Visual Query Builder
- Step 4.3: Visual Filter Builder
- Step 4.4: Bulk Actions Toolbar
- Step 4.5: Smart Suggestions (Trending, ManagedProperty providers)
- Step 4.6: Result Annotations / Tags
- Step 4.7: Audience Targeting
- Step 4.8: Schema Helper (Property Pane Control)
- Step 4.9: Phase 4 Integration Testing

### Phase 5 — Polish & Optimization (7 steps, ~30 tasks)
- Step 5.1: Bundle Optimization
- Step 5.2: Accessibility Audit (WCAG 2.1 AA)
- Step 5.3: Responsive Design
- Step 5.4: Error Handling & Empty States
- Step 5.5: Performance Profiling
- Step 5.6: Documentation
- Step 5.7: Final Validation

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
   // Initialize in onInit: await SPContext.basic(this.context, 'WebPartName');
   // Use anywhere: SPContext.sp.web.lists...

   // Fluent UI v8 — ALWAYS tree-shakable imports
   import { Panel } from '@fluentui/react/lib/Panel';
   // NEVER: import { Panel } from '@fluentui/react';

   // DevExtreme — Lazy load heavy components
   const DataGrid = React.lazy(() => import('devextreme-react/data-grid'));
   ```

2. **Architecture Rules**:
   - Web parts NEVER call PnPjs/Graph directly — always through ISearchDataProvider
   - All inter-webpart communication via Zustand store — no SPFx Dynamic Data
   - AbortController on every search — cancel in-flight requests before new ones
   - Registries freeze after first search execution
   - Services use `SPContext.sp` (no SPFI constructor injection)
   - Use `createSPExtractor` for type-safe field extraction from list items
   - Use `BatchBuilder` for batch SharePoint operations
   - Use `SPContext.logger` for structured logging

3. **Data Rules**:
   - SearchHistory list WILL exceed 5,000 items — ALWAYS filter by Author FIRST in CAML
   - CollapseSpecification fails SILENTLY on non-sortable properties — validate before sending
   - Taxonomy refiners use GP0|#GUID format — must resolve to labels via PnP Taxonomy API
   - Date refiners MUST use FQL range() — NOT raw KQL date comparisons

4. **Path Aliases** (configured in tsconfig.json):
   ```
   @store/*      → src/libraries/spSearchStore/*
   @interfaces/* → src/libraries/spSearchStore/interfaces/*
   @services/*   → src/libraries/spSearchStore/services/*
   @providers/*  → src/libraries/spSearchStore/providers/*
   @registries/* → src/libraries/spSearchStore/registries/*
   @orchestrator/* → src/libraries/spSearchStore/orchestrator/*
   @webparts/*   → src/webparts/*
   ```

---

## Commands

```bash
# Development
npm run serve          # Fast-serve with HMR
npm run serve:legacy   # Gulp serve (fallback)
npm run build          # Development build
npm run build:ship     # Production build

# Testing
npm run test:jest      # Run Jest unit tests

# Production
npm run release        # Clean + bundle + package
npm run stats          # Bundle analysis (ANALYZE=1)
```

---

## Starting Point

To continue development:

1. Read `CLAUDE.md` for full architecture context
2. Check `docs/DEVELOPMENT-PLAN.md` for current status
3. Pick the next incomplete step (suggested: Step 4.2 — Visual Query Builder)
4. Use the `.claude/agents/` specialist prompts for domain-specific guidance

**Suggested first task**: Implement `QueryBuilder.tsx` (Step 4.2.1) — expandable panel below search box with DevExtreme-inspired filter builder UI for constructing KQL queries visually.
