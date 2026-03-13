# SP Search — PnP Modern Search Alignment

This document captures the product-cleanup alignment work done to make the solution feel closer to PnP Modern Search in day-to-day administration, while keeping the platform features added during Sprints 0 to 4.

## Alignment Goals

- predictable starter search page behavior
- no required JSON editing for first use
- familiar concepts: query template, selected properties, refiners, verticals, layouts
- safe defaults for site-scoped search pages

## Where the Solution Now Aligns

### Search Box

| Area | Current SP Search behavior | Alignment intent |
|------|----------------------------|------------------|
| Query input | `{searchTerms}` transformation by default | Same mental model as PnP query template tokens |
| Suggestions | Enabled by default | Similar search-assist behavior |
| Search page handoff | Optional via `searchInNewPage` | Matches common PnP landing-page setup |
| Scope model | Optional scope selector | Comparable admin expectation |

### Search Results

| Area | Current SP Search behavior | Alignment intent |
|------|----------------------------|------------------|
| Query template | `{searchTerms}` starter query | Same default search pattern |
| Search scope | `currentsite` | Common site-search baseline |
| Selected properties | Starter column set plus merged core runtime properties | Mirrors PnP emphasis on explicit returned fields |
| Sorting | Starter sort options configured | Similar out-of-box result tuning |
| Paging | Enabled by default | Common search UX expectation |
| Layout set | `List`, `Compact`, `Grid` by default | Keeps the initial switcher focused |

### Filters

| Area | Current SP Search behavior | Alignment intent |
|------|----------------------------|------------------|
| Starter filters | File type, modified date, author | Mirrors typical PnP search page refiners |
| Author filter | Uses `AuthorOWSUSER` | Correct SharePoint people-filter pairing |
| Cross-filter logic | `AND` by default, `OR` supported | Comparable admin expectation |

### Verticals

| Area | Current SP Search behavior | Alignment intent |
|------|----------------------------|------------------|
| Starter verticals | All, Documents, Pages, Sites | Common first-page structure |
| Per-vertical query | Supported | Same core authoring concept |
| Per-vertical provider | Supported | Extends beyond PnP without breaking the same authoring flow |

## Intentional Differences

These are deliberate product choices, not misalignment.

| Area | Difference | Reason |
|------|------------|--------|
| Search Manager | Saved searches, history, health, insights | Product extension beyond PnP |
| DataGrid | Fullscreen, chooser, export, persisted state | Power-user workflow not covered by basic result layouts |
| Graph People vertical | Native Graph provider and org relationships | Better People experience than SharePoint-only search |
| Scenario presets | General, Documents, Knowledge Base, Policy Search, etc. | Faster deployment and standardization |
| Admin validation | Edit-mode warnings for bad config | Prevents common rollout mistakes |

## Starter Defaults Chosen During Cleanup

### Results

- `searchScope = currentsite`
- `queryTemplate = {searchTerms}`
- `pageSize = 10`
- `showPaging = true`
- `pageRange = 5`
- visible layouts: `list`, `compact`, `grid`
- hidden by default: `card`, `people`, `gallery`

### Selected Properties

- `Title`
- `Author`
- `LastModifiedTime`
- `FileType`
- `Size`
- `Path`
- `SiteName`

### Filters

- `FileType` checkbox
- `LastModifiedTime` date range
- `AuthorOWSUSER` people picker

### Verticals

- `all`
- `documents`
- `pages`
- `sites`

## Cleanup Decisions

### Compact versus Grid

The two views were getting too close. The current product position is:

- `List` = primary reading view
- `Compact` = dense scan view
- `Grid` = power-user table view

`Compact` stays available, but it is no longer treated as meaningfully equal to the DataGrid feature set.

### Removed local grid filter row

The DevExtreme filter row only filtered the current loaded page, not the full result set. That behavior was misleading in a search product, so it was removed instead of being left in as a partial feature.

### View/Edit behavior

Classic SharePoint list forms are no longer embedded inside the modern search experience. View and Edit actions open native SharePoint pages in a new tab.

## Residual Differences to Keep in Mind

- PnP Modern Search relies more heavily on returned/refinable managed-property authoring. This solution now matches that mental model, but still adds provider routing and a richer manager surface.
- The DataGrid remains a page-level result renderer, not a separate analytics or list-management surface.
- Scenario presets intentionally bias toward guided setup instead of total authoring freedom on first load.
