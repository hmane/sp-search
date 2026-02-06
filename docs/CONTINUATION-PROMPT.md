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

### Phase 3: User Features (Partial) ✅
Steps 3.1-3.3 complete — SearchManagerService (CRUD), item-level permissions, Search Manager web part with Saved/Shared/Collections/History tabs.

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

### Step 3.4 — Search Sharing (5 tasks)
```
[ ] 3.4.1 Implement ShareSearchDialog.tsx — tabbed dialog: URL, Email, Teams, Users
[ ] 3.4.2 Share via URL: encode full state in URL params, copy to clipboard with toast
[ ] 3.4.3 Share via Email: mailto: with search description + link + optional top N results
[ ] 3.4.4 Share to Teams: deep link https://teams.microsoft.com/l/chat/0/0?message={encoded}
[ ] 3.4.5 Share to Users: PnP PeoplePicker, create SharedSearch entry, set item-level permissions
```

### Step 3.5 — Search History Auto-Logging (3 tasks)
```
[ ] 3.5.1 Auto-logging in SearchOrchestrator — after successful search, dispatch addToHistory()
[ ] 3.5.2 Clicked item tracking: log { url, title, position, timestamp } to history entry
[ ] 3.5.3 Configurable history cleanup TTL (30/60/90 days) in SearchConfiguration
```

### Step 3.6 — StateId Deep Link Fallback (3 tasks)
```
[ ] 3.6.1 Detect when serialized URL exceeds 2,000 chars
[ ] 3.6.2 Save full state JSON to SearchConfiguration list with ConfigType: StateSnapshot
[ ] 3.6.3 Replace URL with ?sid=<itemId> — on page load, fetch state from list, restore
```

### Step 3.7 — Promoted Results / Best Bets (5 tasks)
```
[ ] 3.7.1 Implement promoted result rule evaluation: match query against rules from SearchConfiguration
[ ] 3.7.2 Implement PromotedResultsBlock.tsx — "Recommended" block above organic results
[ ] 3.7.3 Implement layout-adaptive rendering: card style in Card Layout, row in DataGrid/List/Compact
[ ] 3.7.4 Implement dismissible promoted results (session-only, stored in uiSlice)
[ ] 3.7.5 Implement configurable max promoted results per query (default 3)
```

### Step 3.8 — Active Filter Pill Bar (8 tasks)
```
[ ] 3.8.1 Implement ActiveFilterPillBar.tsx — horizontal strip of dismissible pills
[ ] 3.8.2 Implement pill rendering: {Filter Name}: {Human-Readable Value} x
[ ] 3.8.3 Multi-value filters combined into ONE pill with comma-separated values
[ ] 3.8.4 Pill click dismisses filter via removeRefiner(), re-executes search
[ ] 3.8.5 "Clear All" link at end dispatches clearAllFilters()
[ ] 3.8.6 Human-readable display via IFilterValueFormatter for each field type
[ ] 3.8.7 Animate pill add/remove (Fluent UI motion tokens)
[ ] 3.8.8 Sticky behavior when filter panel is in sidebar layout
```

### Step 3.9 — Recent Searches Suggestion Provider (3 tasks)
```
[ ] 3.9.1 Implement RecentSearchProvider — queries SearchHistory list for current user
[ ] 3.9.2 Implement SuggestionDropdown.tsx in Search Box — dropdown showing grouped suggestions
[ ] 3.9.3 Register RecentSearchProvider in SuggestionProviderRegistry
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

### Phase 4 — Power Features (9 steps, ~40 tasks)
- Step 4.1: Advanced Filter Types (TaxonomyTree, PeoplePicker, Slider, TagBox, Toggle)
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
3. Pick the next incomplete step (suggested: Step 3.4 — Search Sharing)
4. Use the `.claude/agents/` specialist prompts for domain-specific guidance

**Suggested first task**: Implement `ShareSearchDialog.tsx` (Step 3.4.1) — this unlocks the sharing workflow that ties together URL state, list storage, and item-level permissions.
