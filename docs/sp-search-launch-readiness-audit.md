# SP Search Launch-Readiness Audit

**Date:** 2026-05-02
**Scope:** Pre-launch audit covering 6 web parts + 1 library component (SP Search)
**Audience profile:** Any SPFx-capable tenant, self-serve, no author hand-holding
**Spec:** `docs/superpowers/specs/2026-05-02-launch-readiness-audit-design.md`

## Front Matter

### Repo Snapshot
_(populated in Phase 9 — see plan Task 9.3)_

### Verification Snapshot
_(populated in Phase 9 — see plan Task 9.3)_

### Differentiator Priorities
1. Modern UI Quality
2. End-User Productivity
3. Multi-Instance / Multi-Context
4. Admin Experience
5. Observable & Diagnosable

### Reconciliation Summary (March 22 → Today)

| Status | Count | Share |
|--------|------:|------:|
| Closed | 32 | 59% |
| Still-Open | 16 | 30% |
| Changed-Form | 5 | 9% |
| Obsolete | 1 | 2% |
| **Total enumerated** | **54** | **100%** |

Of the 54 findings enumerated from `docs/archive/sp-search-comprehensive-audit-2026-03-22.md`, 32 (59%) are closed in current code with cited commit SHAs from the 2026-03-22 fix sweep (BUG-001/002/003/004/005/007/008/009/011/012, MISS-003/005/006/007/008, INC-002/003/004/005/006/008, SEC-003/005, PERF-001/003, A11Y-001/002/003/006, UX-001/002, ARCH-001). 16 (30%) remain Still-Open and break down as: 6 performance / architecture-maintainability items (PERF-002/004/005/006, ARCH-002/003), 5 accessibility / UX polish items (A11Y-004/005, UX-003/006/007), 4 deferred features (MISS-001/002 = Sprint 4 backlog, SEC-004 saved-search JSON validation, UX-004 by-design vertical filter clear), and 1 documented WONTFIX (BUG-010 setTimeout-vs-queueMicrotask comment in `SearchOrchestrator.ts:124-128`). 5 are Changed-Form — partially addressed but still carry residual risk (MISS-004 show-more inconsistent for Taxonomy/People, INC-001 manual-mode pending-count indicator missing, INC-007 Admin Manager has its own manifest but is still 100% inherited from the user Manager, SEC-002 `allow-scripts` retained intentionally for WOPI, UX-005 manual refresh added but not real-time). 1 is Obsolete: BUG-006 debounce-stale-state (orchestrator now intentionally reads fresh state via `_store.getState()` at execution time, see `SearchOrchestrator.ts:203-204`). The largest open concentration is **Performance + Architecture maintainability** (6 of 16 Still-Open), followed by **Accessibility / UX polish** (5 of 16); both feed the T1 Modern UI Quality and Foundations tracks.

### Reading Guide
- **Effort tiers:** S (≤4h) · M (½–1d) · L (1–3d) · XL (>3d)
- **Priority tiers:** P0 (must ship v1.0) · P1 (should ship v1.0) · P2 (v1.1+) · Defer
- **P0 admission rule:** A finding may only be P0 if it ties to (a) a stated differentiator T1–T5, (b) security, (c) data integrity, (d) a "would prevent install" issue, or (e) a journey Blocker with no documented workaround.
- **Roadmap Matrix (Part 4)** is the executable artifact — open it to pick the next thing to do.

---

## Part 1 — The Two Journeys

### Journey A: Day 1 Admin Install
_(populated in Phase 4 — see plan Tasks 4.1–4.3)_

### Journey B: Day 1 End-User Search
_(populated in Phase 5 — see plan Tasks 5.1–5.3)_

---

## Part 2 — Differentiator Tracks

### T1. Modern UI Quality
_(populated in Phase 6 — see plan Task 6.1)_

### T2. End-User Productivity
_(populated in Phase 6 — see plan Task 6.2)_

### T3. Multi-Instance / Multi-Context
_(populated in Phase 6 — see plan Task 6.3)_

### T4. Admin Experience
_(populated in Phase 6 — see plan Task 6.4)_

### T5. Observable & Diagnosable
_(populated in Phase 6 — see plan Task 6.5)_

---

## Part 3 — Foundations Track
_(populated in Phase 7 — see plan Task 7.3)_

---

## Part 4 — Roadmap Matrix
_(populated in Phase 8 — see plan Task 8.1)_

---

## Part 5 — Recommended Sprint Sequencing
_(populated in Phase 8 — see plan Task 8.2)_

---

## Part 6 — Appendices

### Appendix A — March 22 Audit Reconciliation

**Source:** `docs/archive/sp-search-comprehensive-audit-2026-03-22.md`
**Commits inspected since 2026-03-22:** 26 commits across `feat/spfx-1.22-heft-migration` and `main` (the bulk of the audit-fix work landed on 2026-03-22 itself, after the prior audit was authored).

#### Prior audit count note

The prior audit's two summary statistics disagree with each other and with the per-section enumeration:

| Source | Stated total |
|--------|--------------|
| §1 Executive Summary prose ("12 critical/high + 23 medium + 18 low") | 53 |
| §1 Category table totals (7C + 10H + 26M + 20L) | 63 |
| Per-section heading enumeration (this Appendix) | **54** |

This Appendix uses the per-section enumeration as the source of truth. The 54-finding count is derived from: 12 BUG-NNN (§2) + 8 MISS-NNN (§3) + 8 INC-NNN (§4) + 4 SEC (§5; SEC-001 is a cross-reference to BUG-004 and is not double-counted) + 6 PERF (§6) + 6 A11Y (§7) + 7 UX (§8) + 3 ARCH (§9) = 54. The Per-WebPart Summary tables in §10 and the Priority Fix Matrix in §11 use overlapping aggregations and are not separate findings.

#### Reconciliation table

Cross-references are track-level only at this stage (T1–T5 / Foundations). Phase 8.1 will tighten these to specific Roadmap IDs.

| ID | Title | Original Severity | Status | Evidence | Audit Cross-Ref |
|----|-------|------------------:|--------|----------|-----------------|
| BUG-001 | `operatorBetweenFilters` not watched by orchestrator | Critical | Closed | Fix commit `71c7e7d`. Verified: `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts:96,114,150` now tracks `prevOperatorBetweenFilters`, computes `operatorChanged`, and includes it in the re-search if-condition. | T4 |
| BUG-002 | `queryInputTransformation` not watched by orchestrator | Critical | Closed | Fix commit `d0d1fd3`. Verified: `SearchOrchestrator.ts:97,115,151` adds `prevQueryInputTransformation`, `transformationChanged`, and re-search trigger. | T4 |
| BUG-003 | URL filter restoration abandons pending filters | Critical | Closed | Fix commit `b7efae5`. Verified: `src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts:54` defines `URL_FILTER_RESTORE_TIMEOUT_MS = 5000`; lines 701-712 schedule a `setTimeout` that warns and clears `pendingUrlFilters` after 5s. | T3 |
| BUG-004 | XSS risk via `newPageUrl` property | Critical | Closed | Fix commit `89fbbbc`. Verified: `src/webparts/spSearchBox/components/SpSearchBox.tsx:336-343` validates `newPageUrl` must start with `/`, `https://`, or `http://` before navigation; rejects `javascript:` etc. | Foundations |
| BUG-005 | Multi-context URL prefix race condition | High | Closed | Fix commit `d5bd8be`. Verified: `src/libraries/spSearchStore/store/storeRegistry.ts:217` computes `urlPrefix` at context creation time via `_buildStableUrlPrefix(searchContextId)`; lines 224-234 re-subscribe ALL previously-initialized contexts when the second context is created. | T3 |
| BUG-007 | Null reference risk in SearchBox Manager Panel | High | Closed | Fix commit `e199c80` (URL/UI cleanup). Verified: `SpSearchBox.tsx:879` wraps the entire `<Panel>` block in `enableSearchManager && managerService &&`, removing the dangling Panel shell when service is absent. | T2 |
| BUG-008 | `activeLayoutKey` URL sync race condition | High | Closed | Fix commits `7e0b29a` (component effect) + `e199c80` (URL coercion). Verified: `urlSyncMiddleware.ts:837-863` coerces requested layout to first available + normalizes URL via `replaceState`; `src/webparts/spSearchResults/components/SpSearchResults.tsx:506-510` adds defensive store sync effect. | T1 |
| BUG-009 | Scope round-trip loses `kqlPath` and `resultSourceId` | High | Closed | Fix commit `6f226f2` first added base64 scope JSON encoding; later commit `e199c80` (URL cleanup) removed scope from URL serialization entirely. Scope is now persisted to `localStorage` per `SpSearchBox.tsx:112,144-161,428-432` (cf. closed MISS-005). The original lossy round-trip is impossible. | T3 |
| BUG-011 | Suggestion requests not cancelled on unmount | Medium | Closed | Fix commit `30fc9af`. Verified: `SpSearchBox.tsx:126,196,240-241,284-286` adds `suggestionAbortRef` AbortController, aborts on unmount and on each new request, and ignores stale promise resolutions. | T2 |
| BUG-012 | `shareToUsers` silently drops failed user resolutions | Medium | Closed | Fix commit `f402339`. Verified: `src/libraries/spSearchStore/services/SearchManagerService.ts:1349-1408` returns `{ succeeded: string[]; failed: string[] }` and pushes to `failed` on both empty `user.data.Id` and ensureUser exceptions. | T2 |
| MISS-003 | XLSX export not wired to UI | Medium | Closed | Fix commit `bc56e3c`. Verified: `src/webparts/spSearchResults/components/DataGridContent.tsx:755` lazy-imports `./exportXlsx` and triggers download; toolbar button at line 1105. | T2 |
| MISS-005 | Scope selection not persisted | Medium | Closed | Fix commit `61c9b54`. Verified: `SpSearchBox.tsx:112` defines `SCOPE_STORAGE_KEY`; `:144-161` restores on mount; `:428-432` writes on scope change. | T1 |
| MISS-006 | Clear All filters button not implemented | Medium | Closed | Fix commit `2196c81`. Verified: `src/webparts/spSearchFilters/components/SpSearchFilters.tsx:382-389` defines handler; `:483-484` renders the button when `showClearAll && displayFilters.length > 0`. | T1 |
| MISS-007 | Vertical overflow dropdown on narrow screens | Low | Closed | Verified at HEAD: `src/webparts/spSearchVerticals/components/SpSearchVerticals.tsx:33,129-149,204,220,254` uses Fluent `OverflowSet` to collapse excess tabs into a "More" menu (preexisting in `4b7e370`; survives current branch). | T1 |
| MISS-008 | Search scope configuration UI missing | Medium | Closed | Fix commit `51fb0a4`. Verified: `src/webparts/spSearchBox/SpSearchBoxWebPart.ts:385-` uses `PropertyFieldCollectionData` for `searchScopes` (id/label/kqlPath/resultSourceId fields editable in property pane). | T4 |
| INC-002 | KQL validation UI never displayed | Medium | Closed | Fix commit `60177f1`. Verified: `SpSearchBox.tsx:848-851` renders the validation message in a `role="alert"` div when `!kqlValidation.isValid`; `KqlInput.tsx:215` shows the icon with tooltip. | T2 |
| INC-003 | KQL completion breaks on quoted strings | Medium | Closed | Fix commit `a5ec366`. Verified: `src/webparts/spSearchBox/kql/KqlParser.ts:86-104` `findPropertyDelimiter` is now quote-aware; tracks `inQuote` / `quoteChar` and skips delimiters inside `"..."` or `'...'`. | T2 |
| INC-004 | Collections pagination missing (500 item cap) | Medium | Closed | Fix commit `564e4ce`. Verified: `SearchManagerService.ts:1007-1064` paginates owned items with `Id gt {lastId}` pattern, page size 500, until `batch.length < PAGE_SIZE`. Shared items paginated identically lines 1067-1094. | T2 |
| INC-005 | Base refiner query uses pageSize=1 instead of 0 | Low | Closed | Fix commit `2fdb118`. Verified: `SearchOrchestrator.ts:353` (`pageSize: 0` for base refiner query) and `:582` (vertical count query also uses `pageSize: 0`). | T5 |
| INC-006 | `Store.reset()` doesn't reset AbortController | Low | Closed | Fix commit `ed973a1`. Verified: `src/libraries/spSearchStore/store/createStore.ts:33-37` aborts current `abortController` before resetting state, and resets `abortController: undefined` in the patch (line 45). | T3 |
| INC-008 | ClickedItems JSON can exceed field size limit | Low | Closed | Fix commit `f0ec7c3`. Verified: `SearchManagerService.ts:857-860` defines `MAX_CLICKED_ITEMS = 10` and trims via `splice` before append. | T2 |
| SEC-003 | Collection name not length-validated | Low | Closed | Fix commit `16b387a` (low-priority audit fixes). Verified: `SearchManagerService.ts:1132-1134` rejects names longer than 200 chars in `createCollection`. | Foundations |
| SEC-005 | Teams share URL hardcoded (sovereign cloud failure) | Medium | Closed | Fix commit `b899efd`. Verified: `src/webparts/spSearchManager/components/ShareSearchDialog.tsx:156` uses `getTeamsBaseUrl()` instead of hardcoded `https://teams.microsoft.com`. | Foundations |
| PERF-001 | ActiveFilterPillBar sequential async formatter calls | Medium | Closed | Fix commit `bbf0acf`. Verified: `src/webparts/spSearchResults/components/ActiveFilterPillBar.tsx:163-177` uses `await Promise.all(unresolvedFilters.map(...))` instead of sequential `await` in a loop. | T1 |
| PERF-003 | Schema loaded twice (KQL + Query Builder) | Low | Closed | Fix commit `e199c80` (UI improvements bundle). Verified: `SpSearchBox.tsx:503-506` early-returns from `loadSchema()` when `schemaLoading &#124;&#124; schemaProperties.length > 0`. | T2 |
| A11Y-001 | KQL Input `aria-expanded` hardcoded to false | Medium | Closed | Fix commit `149ffcd`. Verified: `src/webparts/spSearchBox/components/KqlInput.tsx:210` `aria-expanded={!!props.completionsVisible}` reflects dropdown visibility. | Foundations |
| A11Y-002 | Suggestion dropdown missing `aria-activedescendant` | Medium | Closed | Fix commit `149ffcd`. Verified: `src/webparts/spSearchBox/components/SuggestionDropdown.tsx:189` sets `aria-activedescendant={activeIndex >= 0 ? 'suggestion-' + activeIndex : undefined}` and `:213` writes corresponding `id` on each option. | Foundations |
| A11Y-003 | Gallery thumbnails missing `aria-label` | Medium | Closed | Fix commit `149ffcd`. Verified: `src/webparts/spSearchResults/components/GalleryLayout.tsx:109` `aria-label={'View ' + item.title}` on the role="button" thumbnail. | Foundations |
| A11Y-006 | Suggestion remove button no keyboard shortcut | Low | Closed | Fix commit `149ffcd`. Verified: `SuggestionDropdown.tsx:120-125` Delete key invokes `onRemove(activeSuggestion)` when the active suggestion has a `removeAction`. | Foundations |
| UX-001 | Sort dropdown visible on non-sortable layouts | Low | Closed | Fix commit `0522272`. Verified: `SpSearchResults.tsx:793` passes `showSortDropdown && ['list', 'compact', 'grid'].indexOf(activeLayoutKey) >= 0` so People/Gallery/Card hide the sort dropdown. | T1 |
| UX-002 | Empty state message could be smarter | Low | Closed | Fix commit `6ae1b40`. Verified: `SpSearchResults.tsx:402-421` `EmptyState` renders four distinct messages for the (queryText × hasActiveFilters) combinations. | T1 |
| ARCH-001 | Collection identity uses first item's list ID | Medium | Closed | Fix commit `4b7e370` / `8e998e2` era and surviving at HEAD. Verified: `SearchManagerService.ts:294` collection.id is `_hashCollectionName(collectionName)`, not the first list item's `Id`. Deleting any single item no longer invalidates the collection key. | T2 |
| BUG-010 | Vertical layout switch causes one-frame flicker | Medium | Still-Open | Documented WONTFIX. `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts:124-128` has an explicit comment block stating `setTimeout(0)` is REQUIRED — `queueMicrotask()` would re-enter the same subscription call stack and infinite-loop. Documented in commit `4c22bf3`. The flicker remains by design and would require re-architecting the orchestrator subscription to remove. | T1 |
| MISS-001 | Query input transformation not applied | High | Still-Open | Sprint 4 backlog per `CLAUDE.md`. The orchestrator now triggers a re-search when `queryInputTransformation` changes (BUG-002 closed), but the broader concern — that complex transformation patterns are advertised in the property pane without full effect — is not yet fully addressed end-to-end. Reference: `SearchOrchestrator.ts:97,115,151,704-706`. | T4 |
| MISS-002 | `operatorBetweenFilters` not functional | High | Still-Open | Sprint 4 backlog per `CLAUDE.md`. Orchestrator wiring closed (BUG-001), but the actual filter execution path in `SearchService.buildRefinementFilters` may not consistently apply OR semantics across filter groups. Reference: `src/libraries/spSearchStore/services/SearchService.ts` (path called from `SearchOrchestrator.ts:525`). | T4 |
| SEC-004 | SearchState JSON not schema-validated on restore | Low | Still-Open | Verified: `src/webparts/spSearchManager/components/SavedSearchList.tsx:76,163` parses saved-search state with `JSON.parse` followed by a typed cast and try/catch only — no schema check against `IActiveFilter` shape. Tampered JSON could poison the store. | Foundations |
| PERF-002 | KQL completion scans all schema on every keystroke | Medium | Still-Open | Verified: `src/webparts/spSearchBox/kql/KqlCompletionProvider.ts:71-117` linear loop over `schema[]` with `.toLowerCase()` per property per keystroke. No pre-indexed lowercase map. | T2 |
| PERF-004 | Custom `useStoreState` hook verbose shallow comparison | Low | Still-Open | Verified: `src/webparts/spSearchResults/components/SpSearchResults.tsx:188-207` still hand-compares 18 fields. Adding a new store field requires updating the comparator — maintenance risk persists. | T1 |
| PERF-005 | DataGrid color hash runs per-row per-render | Low | Still-Open | Verified: `src/webparts/spSearchResults/components/DataGridContent.tsx:92-104` `getInitialsColor` recomputes per call; no per-name memoization. | T1 |
| PERF-006 | Suggestion `mergeSuggestionsByPriority` creates new Set per call | Low | Still-Open | Verified: `src/webparts/spSearchBox/components/SpSearchBox.tsx:33` allocates `new Set<string>()` on every call (5–6 per keystroke when multiple providers are enabled). | T2 |
| A11Y-004 | Mode toggle buttons use `div` instead of `fieldset` | Medium | Still-Open | Verified: `SpSearchBox.tsx:736` `<div className={styles.kqlModeToggle} role="radiogroup" aria-label="Query input mode">`. ARIA role applied but DOM element is still a div, not semantic `<fieldset>` + `<legend>`. | Foundations |
| A11Y-005 | Scope selector missing `aria-describedby` | Low | Still-Open | Verified: `SpSearchBox.tsx:729` Dropdown only sets `ariaLabel="Search scope"`. No linked description text element. | Foundations |
| UX-003 | Query builder no visual confirmation on apply | Low | Still-Open | Verified: `src/webparts/spSearchBox/components/QueryBuilder.tsx` has no toast/notification on apply; only inline KQL preview updates. The search executes silently. | T2 |
| UX-004 | Vertical tab switching clears all filters | Low (by design) | Still-Open | Verified: `src/libraries/spSearchStore/store/slices/verticalSlice.ts:11` `setVertical` resets `activeFilters: []`. Audit explicitly tagged this as "by design" but the friction remains and is not yet documented in any in-product UI. | T2 |
| UX-006 | Health tab missing user/vertical breakdown | Medium | Still-Open | Verified: `src/webparts/spSearchManager/components/ZeroResultsPanel.tsx:32-64` `aggregateEntries` collapses only by query text + vertical (no per-user view). Admins can't distinguish systemic vs. user-specific issues. | T5 |
| UX-007 | Insights CTR not time-weighted | Low | Still-Open | Verified: `src/webparts/spSearchManager/components/SearchInsightsPanel.tsx` aggregates over the entire window. Comment at `:298` mentions "Daily volume sparkline for trend visibility", but CTR itself is a single window-wide number, not trended week-over-week. | T5 |
| ARCH-002 | Formatter implementation split between store and web part | Low | Still-Open | Verified: formatters exist in BOTH `src/libraries/spSearchStore/formatters/` (registered via `getFilterValueFormatter`) AND `src/webparts/spSearchFilters/formatters/` (`PeopleFilterFormatter.ts`, `BooleanFilterFormatter.ts`, `DateFilterFormatter.ts`, `NumericFilterFormatter.ts`, `TaxonomyFilterFormatter.ts`). Risk of display inconsistencies remains. | Foundations |
| ARCH-003 | Initialization order dependency not enforced | Medium | Still-Open | Verified: `src/libraries/spSearchStore/store/storeRegistry.ts:189` flips `isInitialized = true` and `:191-198` documents the ordering convention in a comment, but nothing programmatically enforces "Results web part calls `initializeSearchContext()` first, then triggers search". | Foundations |
| MISS-004 | Show more/less inconsistent across filter types | Medium | Changed-Form | Partial fix commit `9e26d08` added showMore to TagBox + Dropdown; verified at `src/webparts/spSearchFilters/components/TagBoxFilter.tsx:184`, `DropdownFilter.tsx:146`, `CheckboxFilter.tsx:213`. **Residual gap:** Taxonomy and People filter types still lack a "Show more" affordance — they use type-ahead/tree expand only. | T1 |
| INC-001 | Manual apply mode edge cases | Medium | Changed-Form | Partial fix commit `9c67332` added external-change sync. Verified: `SpSearchFilters.tsx:303-317` syncs pending state with store changes; `:512-523` renders Apply bar. **Residual gap:** No visual count of pending changes (e.g., "3 pending"). The audit's "no visual indicator" point still stands. | T1 |
| INC-007 | Admin Manager is a re-export stub | High | Changed-Form | Fix commit `02c6adc`. Verified: `SpSearchAdminManagerWebPart.ts:16` is `extends SpSearchManagerWebPart`; `onInit` only adds DebugCollector instrumentation. **Residual gap:** behavior is 100% inherited; the only differentiation is preconfigured manifest defaults plus the DebugCollector hook in `onInit`. | T4 |
| SEC-002 | Preview iframe allows scripts | Medium | Changed-Form | Fix commit `c2c1d26` removed `allow-forms`. Verified: `src/webparts/spSearchResults/components/ResultDetailPanel.tsx:282` sandbox is now `allow-scripts allow-same-origin allow-popups`. **Residual:** `allow-scripts` retained intentionally (WOPI requires it; documented in the inline comment), so the original recommendation to drop `allow-scripts` was rejected. | Foundations |
| UX-005 | Zero-result panel not real-time | Low | Changed-Form | Manual `Refresh` button added (`src/webparts/spSearchManager/components/ZeroResultsPanel.tsx:126-128,168`). **Residual:** No subscription to live history writes; admin must click Refresh to see new zero-result entries. | T5 |
| BUG-006 | Debounce timer executes with stale state snapshot | High | Obsolete | Design changed. `SearchOrchestrator.ts:203-204` comment explicitly states `_executeSearch()` reads fresh state via `this._store.getState()` at call time, so changes during the debounce window are *intended* to be captured. The audit framed this as a bug; the resolved design treats fresh-state reads as the contract. | T3 |

### Appendix B — spfx-toolkit Integration Map

This appendix maps newly-shipped `spfx-toolkit` capabilities (since SP Search last integrated) to candidate adoption points in SP Search. It drives Phase 7 / per-track planning, not direct execution — every "Adopt" or "Consider" row should be re-evaluated against the relevant per-track plan before any implementation work begins.

**Toolkit version inspected:** `1.0.0-alpha.1` at `/Users/hemantmane/Development/spfx-toolkit`, commits `1edde9d..920cddb` (2026-01-03 → 2026-04-09; 28 commits inspected via `git log --since='2026-01-01'`). SP Search currently consumes the toolkit via the local file link `"spfx-toolkit": "file:../spfx-toolkit"` in `package.json` — converting that to a published version is itself a Foundations follow-up, not a Phase 2 deliverable.

| Capability | Status | Where it would land in SP Search | Effort | Differentiator | Notes |
|------------|--------|-----------------------------------|:------:|:--------------:|-------|
| **HTML sanitization** (`spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml`, commit `8e9b311`) | Adopt | `src/webparts/spSearchResults/components/documentTitleUtils.ts:155` (`sanitizeSummaryHtml`) — replace local sanitizer used at `ListLayout.tsx:69` `dangerouslySetInnerHTML` | S | Foundations | Toolkit sanitizer has explicit `ALLOWED_TAGS` + `LINK_ATTRIBUTES` allowlists and `isSafeUrl` predicate; replacing the local implementation reduces XSS surface and centralizes the policy. Pairs naturally with the BUG-004 fix sweep called out in the spec §4.4 Foundations track. |
| **Browser storage utilities** (`spfx-toolkit/lib/utilities/browserStorage`, commit `8e9b311`) | Adopt | `SearchManagerService.ts:729,742,764,818` (cleanup-key timestamp), `SchemaService.ts:185,197,206,214` (schema cache), `SpSearchBox.tsx` `SCOPE_STORAGE_KEY` (`:112,144-161,428-432`), `DataGridContent.tsx` (column prefs) | M | Foundations | Centralizes try/catch + availability detection (`isBrowserStorageAvailable`) currently re-implemented inline at 4+ sites. Reduces risk of `QuotaExceededError` regressions called out in `MEMORY.md` PnPjs caching note. |
| **Comments component** (`spfx-toolkit/lib/components/Comments`, commits `523be2a`, `7c798f4`, `434d9ad`, `0aa4dda`, `920cddb`) | Adopt | `src/webparts/spSearchManager/components/ResultAnnotations.tsx` — extend the current tag-only annotation row with a comments thread per saved search / collection item | L | T2 | Toolkit ships hooks (`useComments`, `useCommentInput`, `useCommentSearch`), classic/chat/compact/timeline layouts, @mention search backed by Graph/PeoplePicker, and #link with DocumentLink preview. Directly satisfies a v1.0 productivity differentiator (annotations + collaboration on saved searches / collections). Prerequisite: target SP list with `Comments` field; verify list provisioning script covers it. |
| **ManageAccess panel** (`spfx-toolkit/lib/components/ManageAccess`, commits `2b693b3`, `7cf9186`, `1edde9d`) | Adopt | `src/webparts/spSearchManager/components/ShareSearchDialog.tsx` — replace ad-hoc share-to-users tab with `ManageAccessPanel`, surfacing existing permissions, role assignments, and removal | M | T2 | Current dialog only adds users; cannot view/revoke. ManageAccess covers principal listing + permission level changes + removal in one panel, matching the "personal vs shared library boundaries" deliverable in spec §4.3 T2. Pairs with the Still-Open BUG-012 evidence (`SearchManagerService.ts:1349-1408`) for a unified sharing UX. |
| **CssLoader compatibility aliases** (`spfx-toolkit/lib/utilities/CssLoader`, commits `5ee7cfa`, `7d943a7`) | Consider | `gulpfile.js` `additionalConfiguration` (sp-css-loader exclusion pattern; CLAUDE.md rule #9) | S | Foundations | SP Search currently solves the external-CSS-with-binary-fonts problem at the webpack-config level. Toolkit's CssLoader is an `SPComponentLoader.loadCss`-based runtime loader for Style Library files — solves a different problem. Could simplify admin-managed theming overrides if SP Search exposes one, but no current admin theming surface. Re-assess after T1 (Modern UI Quality) settles its theming story. |
| **FormErrorSummary + FormContext fixes** (`spfx-toolkit/lib/components/spForm`, commits `9ec580e`, `12baab6`, `a5f3ad3`) | Consider | Property pane forms (e.g., `propertyPaneControls/SchemaHelperControl.tsx`) and Search Manager edit dialogs | M | T4 | SP Search property panes currently use plain Fluent UI controls and PnP property controls — not the toolkit `spForm` stack. Adopting `spForm` for edit-mode validation lint (a stated T4 deliverable) is a larger design decision; the recent fixes only matter once `spForm` is the chosen form host. Defer the framework decision to the T4 plan. |
| **ConflictDetector** (`spfx-toolkit/lib/components/ConflictDetector`, commit `df40048`) | Consider | `SearchManagerService.ts` saved-search and collection update paths; `ResultAnnotations.tsx` annotation edits | L | T2 | Saved searches and collections are list items that two admins or the same user across two devices could edit concurrently. ConflictDetector wraps SP ETag checks with a hooks/context API. Worthwhile only if user research shows real concurrent-edit pain; otherwise the optimistic-overwrite default is acceptable for v1.0. Defer scoping to T2 plan. |
| **GroupUsersPicker + GroupViewer fixes** (commits `b43e639`, `a5f3ad3`, `0c8c914`) | Consider | `ShareSearchDialog.tsx` "Share to Users" tab (currently uses PnP PeoplePicker) | S | T2 | If ManageAccess adoption (above) lands first, GroupUsersPicker becomes a sub-decision inside that panel. Re-evaluate only if ManageAccess is rejected. |
| **VersionHistory restyling** (commits `1edde9d`, prior consumption already in `ResultDetailPanel.tsx:7`) | Adopt | `src/webparts/spSearchResults/components/ResultDetailPanel.tsx` — pull the latest VersionHistory styles automatically via existing import | S | T1 | Already consumed; this is a free upgrade once the toolkit version is bumped or re-built locally. SharePoint-native theme alignment supports the T1 Modern UI Quality theming-consistency deliverable. No code change required beyond `npm run build` of the linked toolkit. |
| **UserPersona / userPhotoHelper deduplication** (commits `9c72ec0`, prior consumption in `ResultDetailPanel.tsx:7`, `ListLayout.tsx:4`, `DocumentTitleHoverCard.tsx:8`) | Adopt | Existing UserPersona consumption sites; no new wiring | S | T1 | Performance + correctness improvement to a component already in use across three layouts. Free upgrade on next toolkit rebuild. |
| **NoteHistory dedup / cache-busting** (`spFields/SPTextField/NoteHistory.tsx`, commit `07d1196`) | No Fit | n/a — SP Search does not consume `spFields` or `SPDynamicForm` | — | — | SP Search has no Note field UI; relevant only if a future admin form adopts `SPDynamicForm`. |
| **SPDynamicForm + spFields suite** (`SPLookupField`, `SPTaxonomyField`, etc., commit `13dd811`) | No Fit | n/a — SP Search uses property pane + custom edit dialogs, not toolkit dynamic forms | — | — | Adopting `SPDynamicForm` would be a much larger architectural shift than v1.0 warrants. Tracked under T4 only if a future "admin schema editor" surface materializes. |
| **SPListItemAttachments** (existing) | No Fit | n/a — saved searches and collections do not carry attachments | — | — | No use case in current data model. |

### Appendix C — PnP Modern Search v4 Parity Scorecard

This appendix grades SP Search feature-by-feature against PnP Modern Search v4 (latest stable v4.21.0, released 2026-04-16). It informs **positioning, not a forced parity backlog** — per spec §3 Non-Goals, missing PnP features are deliverables only if they tie to a stated differentiator (T1–T5). PnP v4's documentation surface (per its docs nav: Search Box, Search Results, Search Filters, Search Verticals, Extensibility) is used as the canonical structure to avoid omissions, extending the existing alignment notes in `docs/pnp-modern-search-alignment.md` rather than duplicating them. Grades are: **Better** (SP Search exceeds, citing the SP Search file/feature), **Parity** (equivalent), **Worse** (SP Search has it but inferior, citing the gap), **Missing** (not in SP Search, citing where in PnP it lives).

#### Scorecard

| Area | PnP v4 Feature | Grade | SP Search Equivalent | Notes |
|------|----------------|:-----:|----------------------|-------|
| Search Box | Free-text query input + placeholder + clear | Parity | `src/webparts/spSearchBox/components/SpSearchBox.tsx` | Both ship the basic input, placeholder, and clear-on-Esc. |
| Search Box | Query suggestions (multiple providers, configurable per-group count) | Better | `src/libraries/spSearchStore/providers/{RecentSearchProvider,TrendingQueryProvider,ManagedPropertyProvider,QuerySuggestionProvider,QuickResultsSuggestionProvider}.ts`; suggestion registry in `SuggestionProviderRegistry.ts` | SP Search ships 5 providers vs PnP's 1 built-in (SP Static); registry pattern allows late registration; AbortController cancellation per BUG-011 fix; `aria-activedescendant` per A11Y-002 fix. |
| Search Box | Query input transformation (token-aware template) | Worse | `SpSearchBoxWebPart.ts` exposes `queryInputTransformation` property; `SearchOrchestrator.ts:97,115,151,704-706` watches it for re-search | MISS-001 Still-Open (Sprint 4 backlog): property is surfaced and triggers re-search but the broader transformation pipeline is not fully applied end-to-end (per Appendix A). |
| Search Box | Search-in-new-page (URL fragment or parameter) | Parity | `SpSearchBox.tsx:336-343` validates `newPageUrl` then navigates | XSS-hardened post BUG-004 (`https://`/`http://`/`/` only). |
| Search Box | KQL editor with autocomplete + validation | Better | `src/webparts/spSearchBox/kql/{KqlInput,KqlParser,KqlCompletionProvider}.tsx`; mode toggle in `SpSearchBox.tsx:736` | PnP v4 has no first-class KQL authoring surface; SP Search ships quote-aware parser (INC-003 fix), validation panel (INC-002 fix), and visual Query Builder (`QueryBuilder.tsx`). |
| Search Box | Configurable scopes / result sources | Better | `SpSearchBoxWebPart.ts:385-` `searchScopes` PropertyFieldCollectionData; persisted to localStorage per `SpSearchBox.tsx:112,144-161,428-432` | PnP exposes a single result-source-id property; SP Search ships scope selector + per-scope KQL path + persistence (MISS-005 + MISS-008 closed). |
| Search Results | Data sources (SharePoint Search, Microsoft Search, Azure AI Search, custom) | Worse | `providers/SharePointSearchProvider.ts`, `providers/GraphSearchProvider.ts`; `ISearchDataProvider` interface | SP Search ships 2 built-in providers (SharePoint + Graph); PnP v4 ships 4+ (SP, MS Search, Azure AI Search, custom via `IDataSource`). Per-vertical `dataProviderId` routing (Sprint 3) partially mitigates but Azure AI Search and standalone Microsoft Search are absent. See https://microsoft-search.github.io/pnp-modern-search/usage/search-results/ |
| Search Results | Result layouts catalogue | Better | 6 layouts: `ListLayout.tsx`, `CompactLayout.tsx`, `CardLayout.tsx`, `GalleryLayout.tsx`, `PeopleLayout.tsx`, `DataGridLayout.tsx` (DevExtreme-backed) | PnP ships List/Cards/Tiles/Debug + custom-layout extensibility. SP Search adds DataGrid (column chooser, virtual scroll, CSV+XLSX export, persisted prefs per Sprint 3 + MISS-003 fix) and Graph-backed People layout — both stronger than PnP's Handlebars equivalents on the Modern UI Quality + Admin Experience differentiators. |
| Search Results | Handlebars templating | Missing | n/a — SP Search uses React + cell renderers (`cellRenderers/*`) | PnP v4 layouts are Handlebars + web components (https://microsoft-search.github.io/pnp-modern-search/extensibility/templating/). SP Search's React-component model is a deliberate paradigm difference, not a gap to close — admins extend via `ILayoutDefinition` + cell renderers, not Handlebars. |
| Search Results | Slots (token-bound result fields) | Missing | n/a — SP Search uses `selectedProperties` array + cell renderer mapping | PnP "slots" let admins declare which managed property maps to title/path/preview/etc. without editing the layout. SP Search's columns config is more direct but lacks the slot abstraction layer. https://microsoft-search.github.io/pnp-modern-search/usage/search-results/ |
| Search Results | Tokens (`{searchTerms}`, `{Site.ID}`, `{User.Email}`, etc.) | Parity | `src/libraries/spSearchStore/services/TokenService.ts`; `queryTemplate` defaults to `{searchTerms}` per alignment doc | Token catalogue is comparable per `docs/pnp-modern-search-alignment.md`. |
| Search Results | Sorting (configurable + user-selectable) | Parity | `SpSearchResults.tsx:793` sort dropdown, gated by sortable layouts post UX-001 fix | Sort dropdown hidden for People/Gallery/Card. |
| Search Results | Paging | Parity | `Pagination.tsx`; configurable `pageSize` + `pageRange` per alignment doc | |
| Search Results | Detail / preview panel | Better | `ResultDetailPanel.tsx` with version history, metadata, related docs; iframe sandbox tightened post SEC-002 (Changed-Form) | PnP has no built-in detail panel; admins must hand-author Handlebars hover cards. |
| Search Results | Bulk selection + multi-item actions | Better | `BulkActionsToolbar.tsx`; selection lives in `uiSlice` | PnP has no native bulk actions; admins script via web components. |
| Search Results | Result actions registry (Open/Preview/Share/Pin/Copy/Download/Compare) | Better | `providers/actions/{Open,Preview,Share,Pin,CopyLink,Download,Compare,ExportCsv}Action.ts`; `ActionProviderRegistry` | PnP's action surface is per-template Handlebars; SP Search has a typed registry with 8 built-in actions. |
| Search Results | Promoted results / Best Bets | Better | `PromotedResultsBlock.tsx` (client-side, position #0); admin-defined via SharePoint Query Rules per CLAUDE.md Security Rule #5 | PnP v4 has no client-side promoted results block; relies on invisible server-side ranking. |
| Search Filters | Standard refiners (Checkbox, Date Range, Combo/Dropdown, TagBox, People, Taxonomy Tree) | Parity | `CheckboxFilter.tsx`, `DateRangeFilter.tsx` (FQL `range()` per CLAUDE.md Data Rule #5), `DropdownFilter.tsx`, `TagBoxFilter.tsx`, `PeoplePickerFilter.tsx` (`AuthorOWSUSER`), `TaxonomyTreeFilter.tsx` (GP0\|#GUID resolution) | Catalogue maps 1:1 to PnP's Checkbox/Date Range/Combo/People/Hierarchical templates. Show-more still missing for Taxonomy + People per MISS-004 Changed-Form residual. |
| Search Filters | Date interval (relative ranges: today/week/month/year) | Worse | `DateRangeFilter.tsx` supports custom range only | PnP v4 ships pre-canned interval buckets. https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/ |
| Search Filters | Extra filter types (Slider numeric, Text, Toggle, Visual builder) | Better | `SliderFilter.tsx`, `TextFilter.tsx`, `ToggleFilter.tsx`, `VisualFilterBuilder.tsx` | None of these have PnP v4 equivalents. |
| Search Filters | Operator between filters (AND/OR) | Worse | `SpSearchFiltersWebPart.ts` exposes `operatorBetweenFilters`; orchestrator watches it (BUG-001 closed) | MISS-002 Still-Open: filter execution path may not consistently apply OR semantics across groups (Appendix A row). |
| Search Filters | Manual apply mode + clear-all | Parity | Apply bar `SpSearchFilters.tsx:512-523`; Clear-All `:483-484` post MISS-006 fix | INC-001 Changed-Form residual: no pending-count indicator. |
| Search Filters | Sticky / persisted filter state across vertical switches | Worse | `verticalSlice.ts:11` `setVertical` clears `activeFilters: []` | UX-004 Still-Open by-design choice; PnP keeps filters across verticals when refiner names match. |
| Search Verticals | Tabs with badge counts, per-vertical query/filters config, responsive overflow | Parity | `SpSearchVerticals.tsx` (Fluent `OverflowSet` per MISS-007 closed); `SpSearchVerticalsWebPart.ts` `verticals` collection | |
| Search Verticals | Per-vertical data source / provider routing | Better | Per-vertical `dataProviderId` routes to `SharePointSearchProvider` or `GraphSearchProvider` (Sprint 3 capability per `MEMORY.md`) | PnP v4 verticals share the parent web part's single data source. |
| Extensibility | Typed extension points (custom data sources, suggestions providers, layouts) | Parity | `ISearchDataProvider`, `ISuggestionProvider`, `ILayoutDefinition` with corresponding registries; suggestion registry intentionally NOT frozen per MEMORY.md note | Paradigm difference — SP Search uses TS interfaces + React; PnP v4 uses `IDataSource`/`IExtensibilityLibrary` + Handlebars. https://microsoft-search.github.io/pnp-modern-search/extensibility/ |
| Extensibility | Handlebars helpers, partials, and custom web components | Missing | n/a | PnP v4 layouts are Handlebars + custom HTML web components. SP Search routes the same intent through React composition + cell renderers — deliberate divergence, not a v1.0 deliverable. https://microsoft-search.github.io/pnp-modern-search/extensibility/templating/ |
| Extensibility | Custom query modifiers | Worse | `queryInputTransformation` property surfaced but pipeline incomplete (MISS-001) | PnP ships full IQueryModifier interface. https://microsoft-search.github.io/pnp-modern-search/extensibility/custom-query-modifier/ |
| Extensibility | Custom Adaptive Cards event handlers | Missing | n/a — SP Search does not render Adaptive Cards | https://microsoft-search.github.io/pnp-modern-search/extensibility/custom-event-handlers/ |
| Extensibility | Published compatibility matrix | Worse | No published compatibility matrix; alignment notes touch on it | https://microsoft-search.github.io/pnp-modern-search/extensibility/compatibility-matrix/ |
| Cross-cutting | Search Manager (saved searches, sharing, collections, history, annotations) | Better | `SpSearchManagerWebPart` + `SearchManagerService.ts` (CRUD, item-level perms via `breakRoleInheritance`); `SavedSearchList`, `SearchCollections`, `SearchHistory`, `ShareSearchDialog`, `ResultAnnotations` | PnP v4 has no equivalent — entirely missing on PnP side. Direct T2 differentiator delivery. |
| Cross-cutting | Admin Dashboard (Coverage Stats / Quality Metrics / Health / Insights) | Better | `AdminDashboard.tsx`, `ZeroResultsPanel.tsx`, `SearchInsightsPanel.tsx`, `CoverageStatsService` | PnP has no admin analytics surface. T4 + T5 differentiator delivery. UX-006 / UX-007 are Still-Open polish items. |
| Cross-cutting | Multi-instance isolation via `searchContextId` | Better | `storeRegistry.ts:217` per-context store + `_buildStableUrlPrefix`; `window.__sp_search_context_map__` cross-webpart singleton (CLAUDE.md architecture note); URL namespacing `?ctx1.q=...` | PnP v4 uses Dynamic Data which couples instances; T3 differentiator delivery. |
| Cross-cutting | Debug FAB / DebugPanel (Query/Network/State/Logs/Errors) | Better | `DebugFab.tsx`, `DebugPanel.tsx` | PnP has no in-product diagnostics surface. T5 differentiator delivery. |
| Cross-cutting | Scenario presets (`general`, `documents`, `news`, `people`, `media`, `custom`) | Better | `searchPresets.ts` `SCENARIO_PRESETS`; `_applyScenarioPreset()`; `Search-ScenarioPresets.ps1` | PnP has no preset registry; admins hand-configure each web part. T4 differentiator. |
| Cross-cutting | Theming (Office UI Fabric / Fluent integration) | Parity | Fluent UI v8 throughout; theme-aware CSS variables | Both honour SharePoint section themes. |
| Cross-cutting | Accessibility (WCAG 2.1 AA, keyboard, ARIA) | Worse | A11Y-001/002/003/006 closed; A11Y-004/005 Still-Open | No published WCAG conformance statement on either side; SP Search closed 4 of 6 audit items but has no A11y test pass on file. PnP v4 has no documented A11y baseline either, so the gap is "neither claims AA"; SP Search graded Worse only because it lacks a published conformance statement. Foundations track. |
| Cross-cutting | Mobile responsiveness | Parity | Sprint 3 hardening per `MEMORY.md`: gallery single-column at 399px, overlay backdrop-filter, iOS DataGrid momentum scroll | |
| Cross-cutting | Telemetry (opt-in product signals) | Missing | n/a — no telemetry pipeline | PnP v4 also has no first-class telemetry. SP Search may add opt-in per spec §4.4 Foundations (telemetry plumbing). |
| Cross-cutting | Audience targeting (per-web-part visibility) | Worse | Not implemented as a first-class property | PnP exposes audience targeting on every web part; SP Search relies on SP page-level audience targeting only. |

#### Positioning takeaways

- **SP Search wins on differentiators that PnP v4 doesn't even attempt.** Search Manager (saved/shared/collections/history), Admin Dashboard (Health + Insights), Debug FAB, multi-context isolation via `searchContextId`, scenario presets, KQL editor, and DataGrid layout are all entirely-missing-in-PnP capabilities that anchor T2/T3/T4/T5 differentiators.
- **Filter and refiner catalogue is at parity with one notable gap and one notable strength.** SP Search ships Slider/Text/Toggle and a Visual Filter Builder that PnP doesn't, but lacks PnP's Date Interval (relative-range buckets) — a T1/T4 polish item, easy add.
- **Data source breadth is a real positioning weakness.** PnP v4 ships SharePoint + Microsoft Search + Azure AI Search + custom; SP Search ships SharePoint + Graph only. Per-vertical `dataProviderId` routing softens the gap, but Azure AI Search support is increasingly table stakes for enterprise demos. Decide explicitly whether to invest (T4 deliverable) or document as out-of-scope (Appendix D).
- **Templating paradigm is a deliberate divergence, not a gap to close.** PnP's Handlebars + web components is genuinely more flexible for non-developer admins; SP Search's React + cell renderers + registries is more type-safe and tree-shakable. Document this as a positioning choice in launch materials so prospects don't expect Handlebars.
- **Cross-cutting hygiene (accessibility conformance statement, audience targeting, telemetry) is the area where SP Search is materially behind.** None of these tie to a stated differentiator on their own, but each one fails the "self-serve any tenant" launch bar (spec §3.1) in ways an enterprise procurement review will catch. Foundations track.

#### Sources consulted

- PnP Modern Search docs site root — https://microsoft-search.github.io/pnp-modern-search/ — accessed 2026-05-02
- PnP Modern Search GitHub README (v4.21.0, released 2026-04-16) — https://github.com/microsoft-search/pnp-modern-search — accessed 2026-05-02
- PnP Search Results usage — https://microsoft-search.github.io/pnp-modern-search/usage/search-results/ — accessed 2026-05-02
- PnP Search Filters usage — https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/ — accessed 2026-05-02
- PnP Search Box usage — https://microsoft-search.github.io/pnp-modern-search/usage/search-box/ — accessed 2026-05-02
- PnP Extensibility hub — https://microsoft-search.github.io/pnp-modern-search/extensibility/ — accessed 2026-05-02
- In-repo alignment notes — `docs/pnp-modern-search-alignment.md` (extended by this scorecard, not duplicated)

### Appendix D — Rejected Ideas
_(populated in Phase 8 — see plan Task 8.3)_

### Appendix E — Evidence and Command Log
_(populated in Phase 9 — see plan Task 9.3)_
