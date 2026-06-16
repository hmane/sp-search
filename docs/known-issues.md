# Known Issues & Pending Follow-ups

Tracked limitations and deferred work that are **not derivable from the code** —
captured here so they survive across machines / contributors. Update or remove
entries as they're addressed.

---

## 1. Refiner counts are not cascade-filtered

**Status:** Open · **Area:** Filters / search query construction

A refiner's value counts do **not** reflect the *other* active filters. Example:
with `Hide Inactive Documents = Yes` selected, the `Accounts` refiner still shows
each account's full count (e.g. `1798 (11)`) instead of the active-only count
(`1798 (1)`); the 10 inactive docs are still counted.

**Where:** `SearchService.buildRefinementFilters` (`src/libraries/spSearchStore/services/SearchService.ts`)
builds the `RefinementFilters` from `activeFilters`, and the orchestrator requests
refiners alongside the query. The returned refiner counts aren't narrowing to the
filtered set as expected.

**Care required:** the fix must preserve standard SharePoint behavior where a
multi-select refiner still shows *all* of its own options (so you can add/remove
values within that refiner) while **other** refiners narrow. Don't simply apply
every active filter to every refiner's count query. Needs its own investigation
+ tests.

---

## 2. Version history only works for same-web results

**Status:** Open · **Area:** Result detail panel / spfx-toolkit

The detail-panel "Version history" action errors with *"Unable to load version
history. The item may not exist or you do not have permission"* for any result
that lives in a **different site than the page hosting the search**.

**Root cause:** the spfx-toolkit `VersionHistory` component (`IVersionHistoryProps`)
accepts only `listId` + `itemId` — there is **no `webUrl`/`siteUrl` prop** — and
internally queries the **current page's web** (`SPContext.sp.web` /
`SPContext.webAbsoluteUrl`). Search results are inherently cross-site, so it 404/403s
for anything not in the current web. We also show the button for every result with a
list id (`ResultDetailPanel.tsx`: `hasVersionHistory = !!(ListId && ListItemID)`),
so users hit the error on cross-site docs.

We pass `listId`/`itemId` correctly (`Number(ListItemID)`, both in
`DEFAULT_SELECTED_PROPERTIES`); the limitation is the toolkit, not our usage.

**Fix options:**
- **(A) Make it cross-site capable (real fix):** add an optional `webUrl` to the
  toolkit `VersionHistory` and scope its PnP queries to that web (sibling
  `spfx-toolkit` repo change + rebuild), then pass the result's web URL from the
  panel (request the `SPWebUrl` managed property).
- **(B) Gate the button to same-web results (minimal):** only show "Version history"
  when the item is in the current web — no error, but unavailable for cross-site
  results.

---

## Not an issue (already supported)

- **Per-refiner "Show counts" toggle** already exists: the Filters web part's refiner
  editor (`FiltersCollectionControl`) writes `IFilterConfig.showCount`, respected by
  all filter components (`CheckboxFilter`, `DropdownFilter`, `TagBoxFilter`,
  `TaxonomyTreeFilter`). No new work needed.
