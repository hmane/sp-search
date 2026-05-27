# SP Search — Performance Budgets

> Owned by Foundations track (Found.D7, amended via Task 3.5). CI gate at `scripts/check-bundle-sizes.js`. Adoption deliverables that breach a budget block at the PR review surface; coordinate with track lead before raising a budget.

## Anchor: production minified hashed assets

Budgets are anchored to **production minified bundles** (`heft build --production`), which Heft emits with content-hashed filenames like `sp-search-filters-web-part_<hash>.js`. The gate at `scripts/check-bundle-sizes.js` discovers them via strict regex match `^<stem>_[a-f0-9]{8,}\.js$` — production bundles MUST have a content hash. Unhashed bundles emitted by `npm run package:debug` are intentionally rejected so the gate never silently measures the wrong artefact. SPFx informal "good citizen" guidance is 1–2 MB per web part; current SP Search production bundles all sit comfortably under that bar.

### Amendment history

The original Found.D7 baseline (committed at `38bded6`) was anchored to ~14.75 MB unhashed bundles captured by the audit's Phase 7 inventory. Those numbers turned out to be from a cached non-production build state (likely `npm run package:debug` or stale dev-mode `release/assets/`), not production. Task 3.5 (this amendment) re-anchored the gate against actual production output. The audit's "install-blocking" framing for D7/D12 was based on the stale numbers; the amendment patches `docs/sp-search-launch-readiness-audit.md` accordingly (see Part 3 Foundations Current state and Part 4 Roadmap Matrix Found.D7/D12 rows).

## Per-web-part bundle size budgets (production)

| Web part | Current (bytes) | Budget (1.5×, rounded to 50K) | % of SPFx 2 MB guidance | Notes |
|----------|-----------------|-------------------------------|-------------------------|-------|
| sp-search-filters-web-part.js | 862,498 | 1,300,000 | 43.1% of 2 MB | Largest bundle; carries Filters UI + filter type registry |
| sp-search-results-web-part.js | 714,103 | 1,100,000 | 35.7% of 2 MB | DataGrid lazy-split keeps base under budget |
| sp-search-admin-manager-web-part.js | 592,306 | 900,000 | 29.6% of 2 MB | Insights chunks lazy-loaded |
| sp-search-manager-web-part.js | 591,819 | 900,000 | 29.6% of 2 MB | SearchHistory + Collections lazy-loaded |
| sp-search-box-web-part.js | 546,789 | 850,000 | 27.3% of 2 MB | Suggestion dropdown lazy-loaded |
| sp-search-verticals-web-part.js | 470,158 | 750,000 | 23.5% of 2 MB | Smallest bundle; tabs only |

Bytes column reflects current production output captured by `npm run check:bundles` (mirrored in `docs/performance/bundle-sizes-baseline.json`). Budget column equals 1.5× current rounded up to the nearest 50,000 bytes for headroom.

## Lazy chunk inventory (consumed on demand)

> The chunk sizes below were captured against the prior (non-production) baseline at commit `38bded6`. They are retained for orientation but should be re-captured during a future `webpack-bundle-analyzer` pass before being used as budget anchors.

| Chunk | Size (bytes) | Loaded by |
|-------|--------------|-----------|
| chunk.vendors-fluentui-Dialog | 7,103,488 | SearchManager + AdminManager dialogs |
| chunk.vendors-devextreme-react_core | 3,706,880 | DataGrid Layout (Results) |
| chunk.xlsx_xlsx_mjs | 2,621,440 | DataGrid CSV/XLSX export |
| chunk.vendors-devextreme-react_date-box | 1,124,352 | DateRange filter |
| chunk.spfx-toolkit_PeoplePicker | 2,001,920 | People-picker filter |
| chunk.spfx-toolkit_SearchManager | 545,792 | SearchManager panel |
| chunk.spfx-toolkit_VersionHistory | 450,560 | Detail panel version tab |
| chunk.spfx-toolkit_DataGridContent | 14,336 | DataGrid Layout |

## Enforcement

`scripts/check-bundle-sizes.js` runs after `npm run package` (`heft build --clean --production` + `heft package-solution --production`), discovers `release/assets/sp-search-*-web-part_<hash>.js` (strict — hashed bundles only) via regex, compares against `config/bundle-budgets.json`, and exits non-zero on breach (or on missing/ambiguous matches, including unhashed `npm run package:debug` output). CI wiring lives in the project's Azure DevOps build pipeline (the `npm run check:bundles` step, Found.D8).

Per-PR attribution dashboard at `release/analysis-logs/bundle-sizes.json` captures current vs budget per web part — consumed by reviewers to identify which adopted dependency drove a budget breach.

## Roadmap link

The Found.D12 active reduction sweep was sized against the stale 14.75 MB baseline. With production bundles already well under SPFx informal guidance, the originally-planned XL P0 reduction sweep is no longer warranted; reclassification ships in a follow-up commit that patches `docs/superpowers/plans/` and `docs/sp-search-launch-readiness-audit.md`. Out-of-scope items per audit Foundations Out-of-scope §1: source-mapped production builds.
