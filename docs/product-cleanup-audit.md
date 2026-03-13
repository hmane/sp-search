# SP Search — Product Cleanup Audit

This document records the cleanup pass that aligned fresh-install defaults, web part property contracts, and admin documentation with the current shipped product.

## Scope

The audit focused on four areas:

1. fresh-install web part defaults
2. property-pane and runtime contract alignment
3. PnP-style starter behavior for query, columns, filters, and verticals
4. verification of the current codebase after cleanup

## What Changed

### Starter manifests

Fresh installs now ship with meaningful starter properties instead of near-empty manifests.

| Web part | Cleanup result |
|----------|----------------|
| Search Box | Starter placeholder, query transformation, suggestions, Search Manager, and new-page settings |
| Search Results | Site-scoped starter query, starter selected properties, sort options, paging, and focused layout set |
| Search Filters | Starter refiners for file type, modified date, and author |
| Search Verticals | Starter tabs for all, documents, pages, and sites |
| Search Manager | Real starter properties instead of placeholder description-only config |

### Property contract cleanup

| Area | Result |
|------|--------|
| Author people filter | Standardized on `AuthorOWSUSER` |
| Grid size property | Starter configs and presets aligned to `Size` instead of `FileSize` |
| Config validation | Grid validation now accepts both `Size` and `FileSize` |
| Filter rendering | People/date/toggle filters can render from config even without returned refiner buckets |

### Documentation cleanup

Updated:

- [admin-guide.md](./admin-guide.md)
- [deployment-guide.md](./deployment-guide.md)
- [provisioning-guide.md](./provisioning-guide.md)

Added:

- [pnp-modern-search-alignment.md](./pnp-modern-search-alignment.md)
- [product-cleanup-audit.md](./product-cleanup-audit.md)

## Verification Results

Verification was run after the cleanup changes landed.

| Check | Result |
|-------|--------|
| `npm run type-check` | Pass |
| `npm test -- --runInBand` | Pass |
| `npm run build` | Pass |

### Current automated baseline

- `12/12` test suites passing
- `267/267` tests passing
- build successful on Node `v22.14.0`

## Audit Findings

### No blocking issues found in the cleanup scope

The cleanup changes did not expose any code-level regressions in:

- TypeScript compilation
- unit tests
- SPFx build
- manifest schema compatibility

### Residual risks

These are not failures, but they still need normal release validation on a real SharePoint tenant:

| Area | Why it still needs manual validation |
|------|-------------------------------------|
| Fresh-install starter manifests | Manifest defaults affect new instances, not existing configured web parts |
| Graph People and org chart | Depends on tenant Graph permission approval |
| Search page authoring | Layout, filter, and vertical combinations still need browser-level smoke testing |
| Scenario provisioning script | Requires tenant/site environment validation, not just local build validation |

## Recommended Manual Smoke Pass

Run these checks on a test tenant after deployment:

1. Add new Search Box + Results web parts to a blank page and confirm the starter experience works without manual configuration.
2. Add Filters and confirm file type, modified date, and author appear and refine correctly.
3. Add Verticals and confirm All/Documents/Pages/Sites switch the query as expected.
4. Select the `documents`, `knowledge-base`, and `policy-search` presets and confirm results columns and layouts update correctly.
5. Open the Search Manager and confirm Saved, History, Health, and Insights render after the hidden lists are provisioned.

## Outcome

The product cleanup is complete from a code and documentation standpoint. The solution now has:

- cleaner first-run defaults
- property contracts closer to PnP Modern Search expectations
- updated admin-facing documentation
- a green automated verification baseline
