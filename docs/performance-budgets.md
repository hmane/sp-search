# SP Search — Performance Budgets

> Owned by Foundations track (Found.D7). CI gate at `scripts/check-bundle-sizes.js`. Adoption deliverables that breach a budget block at the PR review surface; coordinate with track lead before raising a budget.

## Per-web-part bundle size budgets (production)

Sizes are unminified `release/assets/sp-search-*-web-part.js` post `heft build --clean --production`. SPFx informal "good citizen" guidance is 1–2 MB per web part; the smallest SP Search bundle is ~4× over that guidance and the largest is ~7× over.

| Web part | Current (bytes) | Budget (1.5×) | Aspirational (50%) | Notes |
|----------|-----------------|---------------|--------------------|-------|
| sp-search-filters-web-part.js | 14,752,081 | 22,128,121 | 7,376,040 | Tree-shake DevExtreme TreeView/FilterBuilder (Found.D12) |
| sp-search-results-web-part.js | 14,503,213 | 21,754,819 | 7,251,606 | Tree-shake DevExtreme DataGrid; defer ResultDetailPanel |
| sp-search-admin-manager-web-part.js | 11,123,315 | 16,684,972 | 5,561,657 | Defer admin-only Insights chunks |
| sp-search-manager-web-part.js | 11,116,954 | 16,675,431 | 5,558,477 | Defer SearchHistory + SearchCollections panels |
| sp-search-box-web-part.js | 8,488,976 | 12,733,464 | 4,244,488 | Defer KQL completion dropdown |
| sp-search-verticals-web-part.js | 7,956,603 | 11,934,904 | 3,978,301 | Smallest bundle; baseline target for v1.1+ |

## Lazy chunk inventory (consumed on demand)

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

`scripts/check-bundle-sizes.js` runs after `heft build --clean --production`, reads `release/assets/sp-search-*-web-part.js` via `fs.statSync`, compares against `config/bundle-budgets.json`, exits non-zero on breach. CI wiring lives in `.github/workflows/build.yml` (Found.D8).

Per-PR attribution dashboard at `release/analysis-logs/bundle-sizes.json` enumerates the top-10 contributing modules per web part — consumed by reviewers to identify which adopted dependency drove a budget breach.

## Roadmap link

Active reduction work belongs to Found.D12 (Tranche 1, P0). Aspirational column targets v1.1+ (post-Sprint 6). Out-of-scope items per audit Foundations Out-of-scope §1: source-mapped production builds.
