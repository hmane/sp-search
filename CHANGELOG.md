# Changelog

All notable changes to SP Search are documented here. Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning follows [SemVer 2.0](https://semver.org/).

## [Unreleased]

This section covers Sprint 5 + Sprint 6 audit Roadmap deliverables landed since v1.0.0-rc.1. All P0 + P1 items are closed; remaining audit work is either P2-deferred-by-design (CI pipeline, design assets, telemetry transport) or out of repo scope. 433 tests pass; all 6 web part bundles within budget.

### Added

#### T1 — Modern UI Quality

- Filters web part collapses to a phone-width drawer below 640 px with Fluent Panel focus trap + Escape-to-close (T1.D1).
- Shared SCSS breakpoint partial `src/styles/breakpoints.scss` consumed by every web part (T1.D2).
- Manager loading skeleton replaces the centered Spinner with a shape-matched Shimmer composition (T1.D3).
- Results renders an idle EmptyState pre-search instead of the 5-row LoadingShimmer (T1.D4).
- Neutral empty-state icon — `Search` / `SearchAndApps` instead of `SearchIssue` warning glyph (T1.D5).
- Detail panel polish: Fluent close button, file-type preview-unavailable card, "Author not indexed" fallback (T1.D6).
- Distinct toolbox icons across all six web parts — no glyph collisions (T1.D7 lite).
- Dimmed-vertical tooltip explains why the tab is unclickable; Apply button renders "Apply N changes" in manual mode (T1.D8).
- `prefers-reduced-motion` honoured universally + new `docs/theming.md` documents dark-mode inheritance (T1.D9).
- Search Box mobile layout: 16 px input font (kills iOS auto-zoom) + min-width 260 px + flex-wrap on the action toolbar (T1.D10).
- Layout-switch scroll preservation via double-RAF restore; layout button tooltips with per-key descriptions (T1.D11).
- New `docs/styleguide.md` consolidating breakpoints, theme tokens, spacing grid, type scale, motion contract, empty-state discipline, loading idiom, button hierarchy, tooltip rules, accessibility floor (T1.D12 partial — visual regression CI deferred to ADO pipeline).

#### T2 — End-User Productivity

- Shared-search recipient notification: badge on tab + MessageBar within 60 s polling window; sender sees "N recipients notified" copy (T2.D1).
- `BulkActionsToolbar` wired across List / Compact / DataGrid layouts; selection persists across layout switch (T2.D2).
- Saved-search JSON schema validation on restore with malformation MessageBar and skip-apply path (T2.D3).
- Owned / Shared-with-me / All toggle on saved searches + collections; per-row "Shared by &lt;Name&gt;" badge on saved searches (T2.D6).
- Detail panel next / previous result navigation with Alt+Left / Alt+Right; "Load next page" at end of page (T2.D7).
- Browser Back / Forward integration via `pushState` for navigational changes, `replaceState` for incremental tweaks (T2.D8).
- Global keyboard shortcuts (`/`, `?`, `Esc`, `Enter`, `Alt+Left/Right`) with `?` help modal listing all six shortcuts (T2.D9).
- Save-flow polish: labelled "Save search" button, context-aware disabled tooltip, scope + layout in the save-dialog preview, auto-save-on-share for unsaved searches with `Untitled search — <date>` generated title (T2.D10 a + b + c).
- Layout-agnostic CSV / XLSX export menu in `ResultToolbar` with "Selection only (N rows)" submenu (T2.D11).
- New `docs/end-user-guide.md` covering Search basics / Open / Save / Share / Owned vs Shared / Collections / Bulk actions / Export / Keyboard shortcuts (T2.D12).
- "Popular Searches" suggestion-group renamed to "Frequent for you" until org-aggregation lands (T2.D13).

#### T3 — Multi-Context Mastery

- Refcounted context dispose: `incrementContextRef` / `decrementContextRef` + microtask-deferred dispose handle the SPFx Modern cross-page navigation race (T3.D1).
- Edit-mode `searchContextId` mismatch banner across all six web parts (T3.D2).
- URL alias collision validator with deterministic disambiguation suffix (T3.D3).
- `searchContextId` field promoted to page 1, group 1, first field on every web part via shared `propertyPaneSearchContextIdField()` helper (T3.D4).
- New `docs/multi-context-guide.md` admin walkthrough — scenarios, authoring checklist, URL namespacing, edit-mode diagnostics, lifecycle, failure-mode table (T3.D5).
- Per-context `urlPrefix` override + `enableUrlSync: false` opt-out via `IInitializeContextOptions` (T3.D6 runtime).
- Per-vertical `dataProviderId` validator with Did-You-Mean + edit-mode MessageBar (T3.D7).
- `DebugPanel.tsx` Multi-Context audit tab listing every context's URL prefix, refcount, init flag, URL sync attachment, registered web parts, store snapshot, with per-row Force-Dispose button (T3.D8).
- `tests/store/lifecycle.test.ts` regression suite — 6 cases covering refcount transitions, dispose-on-zero, deferred-dispose race-window cancellation, state isolation (T3.D9).
- Initialization-order diagnostic catches late-arriving Filters web part; edit-mode MessageBar offers a Retry button (T3.D10).
- New `test-multi-context-tenant-vs-site` provisioning sample demonstrating tenant-wide KB + site-scoped Documents on one page (T3.D11).

#### T4 — Configuration UX

- `SupportsShouldProcess` retrofit on `Setup-SPSearchSite.ps1`, `Provision-SPSearchLists.ps1`, `Map-CrawledProperties.ps1` (plus pre-existing on `Provision-TestData.ps1`); 0 `PSShouldProcess` violations across the four scripts (T4.D1).
- Scenario preset picker promoted to page 1 / group 1; `account-documents` added; People preset always selectable with Graph-warning MessageBar when Graph Org Service is unconfigured (T4.D2).
- `PropertyPaneSchemaHelper` wired into every managed-property field — `selectedProperties`, `compactProperties`, `gridProperties`, `sortableProperties`, `refinementFilters`, `collapseSpecification`, filter `managedProperty` (T4.D3).
- `Get-SeededCoverageProfiles` defaults to tenant-discovered top-N document libraries; `-UseTestData` opt-in restores the legacy hardcoded set; `-MaxSeededLibraries` configurable; AdminMgr empty-state CTA when `coverageProfilesCollection: []` (T4.D4).
- Shared edit-mode validators (`validateCoverageProfileSourceUrls`, `validateExpectedSiteUrls`, `validateManagedPropertyCollection`, `validateRefinementFilterCollection`) wired into Filters / Manager / Admin Manager MessageBar surfaces (T4.D5).
- Admin Manager Path B fork — distinct manifests, distinct toolbox icons (`SearchAndApps` user / `BIDashboard` admin), distinct property pane shapes (user-tab toggles vs admin coverage / health / insights / coverage-profiles), "(Legacy)" copy removed (T4.D6).
- Admin Dashboard depth bundle: Zero-Results 60 s polling (UX-005), per-vertical zero-rate table (UX-006), 7-day rolling CTR sparkline (UX-007) (T4.D7).
- Property-pane field validators for `coverageSourcePageUrl`, `expectedSiteUrls`, `newPageQueryParameter` (T4.D8).
- "Pre-Flight" tab in Admin Manager with 5-row tenant-readiness checklist (Graph permission / hidden lists / SearchHistory ReadSecurity / schema mappings / content source) with green-yellow-red icons + Fix-this links (T4.D9).
- `Deploy-SPSearchSolution.ps1 -ReleaseArtifactUrl` downloads a `.sppkg` from an Azure DevOps / GitHub Releases URL (T4.D10).
- Context-sensitive `?`-style help links on every property pane group across all six web parts, linking to deep anchors in `docs/admin-guide.md`; `setPropertyPaneHelpBaseUrl` override for tenants mirroring docs internally (T4.D11).
- Cross-web-part scenario preset propagation via `recordPresetSuggestion` registry; Filters web part offers Apply / Dismiss MessageBar when Results selects a preset with `filterSuggestions` (T4.D12).

#### T5 — Debug + Observability

- Cross-bundle singleton DebugFab on every user-facing web part via `DebugFabHost` + window-backed owner-claim flag (T5.D1).
- `DebugCollector` Network tab + per-call timing + 50-row in-memory buffer (T5.D2).
- `DebugPanel` tab-registration API (`registerDebugTab` / `getRegisteredDebugTabs`) consumed by T3.D8 Multi-Context tab (T5.D4).
- Central `spLog` shim with PII redaction (queryText / userEmail / userId) + production-gate (warn / error exempt) (T5.D6).

#### Foundations

- Stale-docs sweep: admin-guide Manager defaults corrected to match post-T4.D6 manifests; legacy gulp / SPFx 1.21 references verified absent from `CLAUDE.md`, `docs/deployment-guide.md`, `docs/admin-guide.md`, `docs/provisioning-guide.md` (Found.D5).

#### Audit Appendix A closures

- MISS-001 — `queryInputTransformation` end-to-end token resolution.
- MISS-002 — `operatorBetweenFilters` end-to-end filter execution.

### Changed

- `setLayout` no longer clears `bulkSelection` — selection persists across layout switches (T2.D2 acceptance signal).
- Schema service cache reused across all managed-property fields so a single tenant fetch services every validator + dropdown (T4.D3 substrate).
- `disposeStore` JSDoc rewritten — documents the refcount as the preferred lifecycle path, direct calls only for tests + admin Force-Dispose.

### Fixed

- Silent leak on SharePoint SPA navigation: `disposeStore` is now wired to every web part's `onDispose` via refcount (T3.D1) — `window.__sp_search_context_map__.size` returns to 0 after last unmount.
- Filters-late-registration silently dropped URL-deep-linked filter values: T3.D10 init-order diagnostic surfaces the issue with Retry recovery.
- People vertical silently routing to SharePoint Search on `dataProviderId` typo: T3.D7 edit-mode validator catches it.
- Filter URL alias collisions silently dropping values on deep-link round-trip: T3.D3 disambiguation suffix.

### Security

- T2.D3 — Saved-search JSON schema validation closes SEC-004 (Still-Open at v1.0.0-rc.1).

### Removed

- `.github/` directory removed entirely — the GitHub Actions CI workflow (`build.yml`), release workflow (`release.yml`), Dependabot config (`dependabot.yml`), and GitHub issue/PR templates added in Found.D8. The build/test/bundle gates and `.sppkg` packaging run via the project's Azure DevOps pipeline; dependency review is manual; bugs and PRs use ADO work items.

### Deferred (audit-acknowledged, not shipped in this branch)

- T1.D7 custom illustrative SVG icons — design assets, not code. Audit signal partially met via T1.D7 lite (six distinct existing Fluent icons).
- T1.D12 visual regression CI pipeline — ADO-side, not source.
- T5.D8 telemetry transport + T5.D9 telemetry-derived insights — infrastructure not in v1.0 scope.

### Metadata carried forward from v1.0.0-rc.1

- `package.json:version` aligned to `1.0.0` from generator default `0.0.1`; `config/package-solution.json:solution.version` aligned to `1.0.0.0` (lockstep — Found.D11).
- `solution.developer.mpnId` cleared from `Undefined-1.21.1` to empty string (Found.D11). Populate with a real Partner Center MPN ID once one is registered.
- `solution.developer.websiteUrl / privacyUrl / termsOfUseUrl` populated with canonical project URLs (Found.D11).
- CI/release tooling standardized on Azure DevOps. `docs/release-policy.md`, `docs/performance-budgets.md`, `docs/accessibility.md`, `CONTRIBUTING.md`, `README.md`, and `docs/release-runs/v1.0.0-rc.1.md` updated to reference the ADO pipeline.

## [1.0.0-rc.1] - 2026-05-DD

### Added

- SPFx 1.22 / Heft build pipeline (Foundations Found.D2 — squash-merge of 91-commit feat/spfx-1.22-heft-migration branch).
- Per-web-part bundle size budgets and CI breach gate (`config/bundle-budgets.json`, `scripts/check-bundle-sizes.js` — Found.D7).
- Heft Jest harness via `@rushstack/heft-jest-plugin` shared config; `tests/store/lifecycle.test.ts` smoke trail-marker (Found.D13).
- Top-level `README.md`, `CHANGELOG.md`, `CONTRIBUTING.md`, `docs/release-policy.md`, `docs/release-smoke-checklist.md` (Found.D2/D5/D8).
- Scenario presets for `general`, `documents`, `news`, `people`, `media`, `custom`, `knowledgeBase`, `hubSearch`, `policySearch` (Sprint 3 — `searchPresets.ts:64-384`).
- DataGrid layout with admin-configured columns, cell renderers, filter row, column chooser, virtual scrolling, CSV + XLSX export, localStorage column preferences (Sprint 3).
- Graph-backed People vertical via `GraphSearchProvider` with presence batch (Sprint 3).
- Analytics feedback loop: Health tab (zero-result queries) + Insights tab (top queries / CTR / daily volume) (Sprint 3).

### Changed

- Build pipeline migrated from gulp to Heft (`a5f28c1`); SPFx 1.21.1 → 1.22.2; spfx-toolkit type alignment (`77adef7`).
- `package.json:type-check` script now invokes `tsc --noEmit -p tsconfig.json` directly (Found.D3).
- Gallery layout collapses to single-column at 399px viewport (Sprint 3 mobile hardening).
- Admin Manager toggles (enableSavedSearches/Shared/Collections/History) ship `false` by default per `SpSearchManagerWebPart.manifest.json` (admin must opt in per tab); admin-guide updated to match (Found.D5).

### Fixed

- BUG-001..BUG-012 closures from the 2026-03-22 audit reconciliation pass (see `docs/sp-search-launch-readiness-audit.md` Appendix A).
- BUG-004 (XSS via `newPageUrl`): closed via `https?://` / `/` allowlist on `SpSearchBox.tsx:358`; remaining 7 unhardened sites consolidated into `safeNavigate` helper (Found.D4 follow-up).
- `pnpPropertyControlsFix.ts` ESLint `no-use-before-define` blocker that halted `npm run package` (Found.D1).
- `SearchHistory` Author-first CAML predicate to prevent threshold throttling on >5,000-item lists.
- PnPjs caching `QuotaExceededError` handled via inline retry + outer catch.

### Security

- SEC-003 (collection name length validation) closed.
- SEC-005 (Teams URL sovereign-cloud handling) closed.
- A11Y-001/002/003/006 (KQL ARIA + gallery aria-label + suggestion keyboard shortcut) closed.
