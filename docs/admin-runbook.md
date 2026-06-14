# SP Search — Admin Runbook

Symptom → diagnosis → resolution playbook for SharePoint admins running SP Search in production. Use this when something doesn't work. For initial setup, see [deployment-guide.md](./deployment-guide.md). For property-pane configuration, see [admin-guide.md](./admin-guide.md).

## Quick reference

| Class of issue | Jump to |
|---|---|
| Search returns nothing / empty results | [Empty results](#empty-results) |
| Filters missing, broken, or wrong | [Filter problems](#filter-problems) |
| Verticals tab is dimmed / wrong count / doesn't switch | [Vertical tab problems](#vertical-tab-problems) |
| People search / People filter empty | [Graph / People problems](#graph--people-problems) |
| Web part doesn't render on the page | [Web part doesn't load](#web-part-doesnt-load) |
| Result click does the wrong thing (wrong tab, blank modal, blocked) | [Result link / preview problems](#result-link--preview-problems) |
| Save / Share / Collections / History broken | [Manager problems](#manager-problems) |
| Admin Manager tabs empty / Dashboard blank | [Admin Manager problems](#admin-manager-problems) |
| URL doesn't share or restore search state | [URL sync problems](#url-sync-problems) |
| Web part hidden for some users but not others | [Audience targeting problems](#audience-targeting-problems) |
| Performance / slow load | [Performance](#performance) |

If none of those match, run the Pre-Flight diagnostic first — it catches the most common install-level issues. See [Pre-Flight](#pre-flight-self-test).

## Pre-Flight self-test

The **Admin Manager → Pre-Flight** tab runs an automated readiness check. Open it before deep-diving any other symptom — it surfaces these failure modes one click:

| Check | What it verifies | Resolution if it fails |
|---|---|---|
| Graph `People.Read` / `User.Read` permissions | Tenant has approved the Graph permissions declared by the .sppkg. `People.Read` powers the People vertical; `User.Read` powers audience targeting via `/me/memberOf`. | SharePoint admin centre → Advanced → API access → approve the pending requests |
| Core hidden lists exist | `SearchSavedQueries`, `SearchHistory`, `SearchCollections` are present on the site | Re-run `Deploy-SPSearchSolution.ps1 -ProvisionSite` or `Setup-SPSearchSite.ps1` |
| SearchHistory item-level security | `ReadSecurity = 2`, `WriteSecurity = 2` (users see only their own rows) | Re-run `Deploy-SPSearchSolution.ps1`; the script sets this via CSOM after template apply |
| SearchHistory runtime fields | `QueryText`, `QueryHash`, `Vertical`, `SearchPageUrl`, `SearchState`, `UseCount`, `ResultCount`, `IsZeroResult`, `ClickedItems`, and `SearchTimestamp` exist | Re-run `Provision-SPSearchLists.ps1` to repair missing fields, or re-run `Deploy-SPSearchSolution.ps1 -ProvisionSite` with the current template |
| Core hidden-list indexes | `Author` indexes exist on `SearchSavedQueries`, `SearchHistory`, and `SearchCollections`; high-volume date/hash/name columns are indexed | Re-run `Provision-SPSearchLists.ps1` before the lists exceed 5,000 items |
| Schema mappings | Crawled→managed property mappings the Results web part needs | Run `Map-CrawledProperties.ps1` against the target site |
| Content source | A SharePoint content source is indexing the target site | Tenant search admin centre — check crawl status |

If Pre-Flight passes and you still have a problem, jump to the matching symptom section below.

## Empty results

**Symptom:** Searches consistently return zero results, or the empty state shows on every query.

### Diagnose

1. Try the same query in `https://<tenant>.sharepoint.com/_layouts/15/osssearchresults.aspx?k=<query>` (SharePoint's stock search page). If it also returns nothing → indexing or content-source issue, not SP Search.
2. Check the Results web part's **search scope** in the property pane. The default is `currentsite` — if the content lives on a *different* site, the scope needs to be `tenant`, `hub`, or `custom` with the right path.
3. Check **vertical query templates**. If the active vertical has a `queryTemplate` like `{searchTerms} IsDocument:1` but the user's looking for pages, no results will return.
4. Open `?debug=1` on the page → **Network tab** in the DebugFab → look at the actual KQL the orchestrator sent. Often the template + transformation produces a query that's narrower than intended.

### Resolution

| Cause | Fix |
|---|---|
| Content not indexed | Wait 15–60 min after upload; check tenant search admin centre |
| Wrong scope | Property pane → Results → Data → Search scope |
| Restrictive vertical template | Property pane → Verticals → edit the active vertical → simplify the query template |
| `queryInputTransformation` overly narrow | Property pane → Search Box → Search → check transformation; default is `{searchTerms}` |
| Filters silently applied via URL | Click **Clear all** in the filter pill bar (or strip `&ft=...` params from the URL) |
| User has no permission on indexed content | Search trims results by ACL; non-admins won't see admin-only content |

## Filter problems

**Symptom:** Filter sidebar empty, filter has no values, removed filter values come back, or applying a filter has no effect.

### Diagnose

1. **No filter sidebar at all** — confirm the Filters web part is on the page and its `searchContextId` matches Results.
2. **Filter shows but has no values** — managed property isn't refinable. Site Settings → Search Schema → find the property → verify the **Refinable** flag is set. If not, add a mapping (`RefinableString00`, etc.) via `Map-CrawledProperties.ps1`.
3. **People filter returns nothing** — admin used `Author` instead of `AuthorOWSUSER`. The Filters web part normalises legacy `Author` to `AuthorOWSUSER` at runtime, but if a different people property is misspelled, it'll silently fail.
4. **Date refiner doesn't filter** — date refiners use FQL `range()`, not raw KQL date comparisons. If the managed property isn't a `DateTime` type, the FQL token is invalid and SharePoint silently returns the unfiltered set.
5. **Taxonomy refiner shows GUIDs not labels** — Taxonomy tokens (`GP0|#<guid>`) need to resolve to term labels via the PnP Taxonomy API. If the user doesn't have term-store read permission, labels won't resolve.

### Resolution

| Cause | Fix |
|---|---|
| Property not refinable | Add managed property mapping; mark Refinable. Allow ~15 min for re-crawl. |
| People filter using `Author` | Switch to `AuthorOWSUSER` in the Filters property pane (auto-normalised, so this is cosmetic only) |
| Date refiner on wrong type | Use only on `DateTime`-typed managed properties (`LastModifiedTime`, `Created`, etc.) |
| Taxonomy labels not resolving | Grant the audience read access to the term store (Term Store Management) |
| Filter applies but UI doesn't update | Hard-refresh (Cmd+Shift+R) — likely a stale bundle from prior deploy |

## Vertical tab problems

**Symptom:** Vertical tabs greyed out, counts wrong, or clicking a tab doesn't switch the result set.

### Diagnose

1. **All tabs dimmed** — current query returns zero results in every vertical. Try a broader query or a different scope.
2. **One specific tab dimmed and unclickable** — that vertical's `queryTemplate` produced zero hits for this query. Hover the tab for the tooltip: "No results in this vertical for the current query."
3. **Tab count badge wrong** — vertical count comes from a parallel count query. If the count query fails (network error, abort), the badge falls back to 0. Check `?debug=1` → Network tab.
4. **Clicking tab does nothing** — `searchContextId` mismatch between Verticals and Results. The Verticals web part dispatches to the store, but if Results is on a different context, it never reacts.
5. **Link-mode vertical (configured as external link) still navigates when dimmed** — fixed in current build (`aria-disabled` + `href` dropped). If you see this, re-publish the page after upgrading the .sppkg.

### Resolution

| Cause | Fix |
|---|---|
| Cross-vertical count query failed | Check Graph / SharePoint search service availability; reload the page |
| Context mismatch | Property pane → all connected search web parts → set the same `searchContextId` |
| Vertical query too narrow | Property pane → Verticals → edit query template |
| Old bundle in browser cache | Hard-refresh; if still stale, clear SharePoint CDN cache or wait for CDN TTL |

## Graph / People problems

**Symptom:** People vertical empty, People filter returns nothing, org chart not rendering.

### Diagnose

1. Open SharePoint admin centre → Advanced → **API access**. Find pending requests for **Microsoft Graph**. If you see `People.Read` pending, that's the blocker.
2. Open `?debug=1` → Network tab → trigger a People search. If you see `403 Forbidden` from `/search/query`, the permission isn't granted to the SPFx app's identity.
3. **People vertical configured but routes to SharePoint Search** — the vertical's `dataProviderId` is misconfigured. Per-vertical override should be `dataProviderId = graph-people` (not `graph`, not `sharepoint`).
4. **Org chart UI hides itself** — graceful fallback when `User.Read.All` isn't granted. The People card still renders, just without the manager/direct-reports section.

### Resolution

| Cause | Fix |
|---|---|
| `People.Read` not approved | SharePoint admin centre → Advanced → API access → approve. Wait 5–10 min for the change to propagate. |
| Wrong `dataProviderId` | Property pane → Verticals → people vertical → set `dataProviderId = graph-people` |
| `User.Read.All` not approved | Org-chart feature degrades silently; non-blocking. Approve when convenient for full org-chart UI. |
| Presence dots not showing | Presence requires `Presence.Read.All` (Graph). Optional — falls back to no dot. |

## Web part doesn't load

**Symptom:** Page shows blank space where a search web part should be, or web part renders with no content / error.

### Diagnose

1. Open browser console. Look for errors prefixed `[SP Search]`. Common ones:
   - `Cannot read properties of undefined (reading 'getState')` → store not yet initialised (race; fixed in current build with `_store` guards in all six web parts).
   - `Failed to fetch` / `403` → permission or auth issue.
   - `Refused to display ... in a frame` → CSP / X-Frame-Options on the iframe target.
2. Open `?debug=1` → DebugFab → check the **Multi-Context** tab. It lists every `searchContextId` on the page, refcount, init status, and registered web parts. If a web part is missing from the list, it never registered with the store.
3. In edit mode, look for an inline **MessageBar** banner above the web part. It announces `searchContextId` mismatch and init-order issues that the runtime detected.

### Resolution

| Cause | Fix |
|---|---|
| Web part renders before `onInit` completes | All six web parts now guard with `if (!this._store) return;` in `render()`. If you still see this on an upgraded build → hard-refresh + clear node_modules/.cache and re-deploy. |
| `searchContextId` mismatch | Edit-mode banner tells you which web part is on the wrong context. Fix in property pane → first field on every web part. |
| Filters web part renders late (after Results' first search) | Init-order diagnostic surfaces this as an edit-mode MessageBar. Re-arrange page sections so Filters loads before Results (Section 1 = Box, Section 2 = Verticals, Section 3 = Results + Filters as columns). |
| Audience-targeted web part hidden for current user | See [Audience targeting problems](#audience-targeting-problems) below |
| Old bundle cached | Hard-refresh (Cmd+Shift+R) + `npm run clean:cache` on the dev box if seeing locally |

## Result link / preview problems

**Symptom:** Clicking a result does the wrong thing — opens in wrong tab, shows "blocked by Chrome", or the modal preview is broken.

### Diagnose

1. Open browser DevTools → Elements → inspect the `<a>` element on the result title.
   - `data-interception="off"` should be present. If not, the page is on an old bundle.
   - `target="_blank"` should be present when `clickTarget=panel` or `newTab`. If missing, `clickTarget` is set to `sameTab` or `sidePanel`.
2. Click a PDF. Should open the **preview popup Modal**. If it navigates the current tab → SharePoint Modern's SPA router hijacked the click (missing `data-interception="off"`, fixed in current build).
3. Modal opens but shows "**This page has been blocked by Chrome**" → old bundle without the `<embed>`-for-PDFs change. The current build renders PDFs via `<embed type="application/pdf">` (browser-native viewer) and only uses sandboxed `<iframe>` for Office docs.
4. Modal opens but blank → tenant blocks WopiFrame in iframes via `X-Frame-Options: DENY`. Office Online preview requires SAMEORIGIN. Confirm with browser DevTools → Network → look for the WopiFrame request's response headers.

### Resolution

| Cause | Fix |
|---|---|
| Current tab navigates instead of new tab / Modal | Old bundle. Re-deploy + hard-refresh. Verify `<a>` has `data-interception="off"` attribute. |
| PDF shows "blocked by Chrome" | Old bundle. Re-deploy; current build uses `<embed>` for PDFs. |
| WopiFrame iframe blank | Tenant CSP / X-Frame-Options; coordinate with SharePoint admin to allow same-origin framing of `_layouts/15/WopiFrame.aspx`. As a workaround, switch `clickTarget` to `newTab` in the Results web part. |
| Side-panel mode shows nothing when clicking | `enablePreviewPanel = false` in property pane. Set it to `true` for `clickTarget = sidePanel` to work. |
| Detail panel "Next" jumps slowly at page boundary | Current build auto-advances to the next page's first item; if it doesn't, old bundle. |

## Manager problems

**Symptom:** Save / Share doesn't work, shared search recipients don't see it, history empty.

### Diagnose

1. **Save button disabled** — hover for the tooltip. It tells you which precondition isn't met (typically: empty query + no filters applied).
2. **Save succeeds but search doesn't appear in the list** — `SearchSavedQueries` hidden list either missing or permissions broken. Verify in `https://<site>/Lists/SearchSavedQueries/AllItems.aspx`.
3. **Share succeeds but recipient sees nothing** — item-level security broken. The Share action calls `breakRoleInheritance()` + `addRoleAssignment(<recipient>)` on the saved-search list item. If the user doesn't have permission to break role inheritance on the list, the item permissions don't update. List should have inheritance broken globally (see Pre-Flight).
4. **History empty** — `SearchHistory` list missing OR one of its runtime fields is missing. The runtime writes a row per query; if the write fails, the history stays empty. Check browser console for `[SP Search] SearchManagerService.logSearch failed` or schema-mismatch warnings.
5. **Notification badge doesn't clear** — the dismiss handler writes `Acknowledged = true` on the share row; if write fails (permission issue), the badge persists. Check Network tab for `403` on the PATCH.

### Resolution

| Cause | Fix |
|---|---|
| Hidden list missing | Re-run `Deploy-SPSearchSolution.ps1 -ProvisionSite` or `Setup-SPSearchSite.ps1` |
| `The field or property 'QueryText' does not exist` or another SearchHistory field is missing | Re-run `Provision-SPSearchLists.ps1 -SiteUrl <site> -ClientId <app-id> -Force` to add missing fields without deleting history |
| `IsZeroResult` field missing on SearchHistory | Re-run `Provision-SPSearchLists.ps1 -SiteUrl <site> -ClientId <app-id> -Force`, or add manually via `Add-PnPField -List SearchHistory -DisplayName "IsZeroResult" -InternalName "IsZeroResult" -Type Boolean` |
| Item-level security broken on SearchHistory | Re-run `Deploy-SPSearchSolution.ps1`; the script sets `ReadSecurity = 2`, `WriteSecurity = 2` via CSOM |
| Share recipient sees no shared search | Verify recipient has `Read` on the list (`Get-PnPListPermissions`); the per-item grant relies on the list-level baseline |
| User isn't a Member of the site | Share recipients need to be at least Members to receive shared search rows |

## Admin Manager problems

**Symptom:** Admin Manager web part shows access-denied, tabs missing, or Dashboard sections empty.

### Diagnose

1. **Access denied panel** — current user isn't `ManageWeb` (Owner/Admin) on the site. By design. Only Owners/Admins see this web part.
2. **No Pre-Flight tab** — old bundle. Pre-Flight is admin-variant-only; if it's missing on a deployed Admin Manager, re-deploy.
3. **Dashboard → Content Coverage shows "No coverage profiles configured"** — `coverageProfilesCollection` is empty. Either run `Setup-SPSearchSite.ps1` to seed the top-5 libraries automatically, or configure profiles by hand in the property pane → Monitoring → Coverage profiles.
4. **Dashboard → Search Quality empty** — needs accumulated history. CTR + zero-result-rate are computed from `SearchHistory`. New sites with no history show empty cards; populate by running real queries for a few days.
5. **Health tab empty** — no zero-result queries in history yet. Either history is genuinely zero-result-free, or `IsZeroResult` field is missing (see Manager problems above).
6. **Insights tab empty** — same root cause as Health (no history rows / missing field).

### Resolution

| Cause | Fix |
|---|---|
| Not Owner/Admin | Expected behavior. The Admin Manager web part is gated by `ManageWeb`. |
| Coverage profiles empty | Run `Setup-SPSearchSite.ps1` (auto-seeds) OR configure manually in property pane |
| Quality metrics blank on a fresh site | Accumulate history by running real queries for ~1 week |
| Health / Insights blank | Verify `IsZeroResult` field exists on SearchHistory (see Pre-Flight) |
| Pre-Flight missing | Re-deploy current build; old bundles didn't have Pre-Flight |

## URL sync problems

**Symptom:** Shared search URL doesn't restore filters, or browser Back/Forward doesn't move through search history.

### Diagnose

1. Look at the actual URL after applying a filter. You should see params like `?q=annual%20report&v=documents&ft=docx,pptx`. If the URL doesn't change → URL sync middleware not attached (init-order issue).
2. Open `?debug=1` → Multi-Context tab → check "URL sync attached" column. Should be `true` for the active context.
3. **Filter values garbled in URL** — taxonomy/people filter values include special chars (`|`, `#`). The URL encoder uses `encodeURIComponent`; if a downstream consumer (mail client, Teams) re-encodes, you'll get double-encoded URLs.
4. **Two contexts on one page conflict** — both write to `?q=`. Multi-context pages must set an explicit `urlPrefix` on each context (e.g. `ctx1`, `ctx2`) → params become `?ctx1.q=...&ctx2.q=...`.

### Resolution

| Cause | Fix |
|---|---|
| URL sync not attached | Hard-refresh; if persistent, check init-order banner in edit mode |
| Multi-context URL clash | Set `urlPrefix` via `initializeSearchContext(ctxId, ctx, { urlPrefix: 'ctx1' })` — see [multi-context-guide.md](./multi-context-guide.md) |
| Browser Back/Forward not working | Current build uses `pushState` for navigational changes; check `popstate` listener in DebugFab |
| URL too long to share | State-ID fallback engages automatically for complex state; URL becomes `?sid=<id>` and the rest is stored in the `SearchSavedQueries` list |

## Audience targeting problems

**Symptom:** Web part visible to some users, hidden for others, or hidden for everyone.

### Diagnose

1. Check the web part's property pane → **Audience targeting** group → look at the configured Microsoft Entra group object IDs. Empty = visible to everyone; non-empty = restricted.
2. Per-vertical / per-refiner / per-promoted-result audience targeting works the same way — each can carry its own audience list.
3. **Hidden for everyone** — audience group IDs are configured, but Graph couldn't resolve the current user's group membership (failed `/me/memberOf` call). Default is fail-closed → nobody sees the surface until membership resolves.
4. Open `?debug=1` → check the store's `currentUserGroups` value. If `[]`, Graph call failed or hasn't completed.

### Resolution

| Cause | Fix |
|---|---|
| `User.Read` Graph permission not granted | SP admin centre → API access → approve. The least-privilege scope for `/me/memberOf`. |
| Group object ID typo | Open the Microsoft Entra group, copy the **Object ID** (GUID), paste into the audience field. Display name doesn't work. |
| Want visible to all again | Clear the audience field in the property pane. Empty = visible. |

## Performance

**Symptom:** Slow first paint, slow result render, page feels heavy.

### Diagnose

1. Open browser DevTools → Network → reload. Filter by JS. Bundle sizes are checked in CI via `scripts/check-bundle-sizes.js` — if any bundle is unexpectedly large, the gate would have failed pre-deploy.
2. Heavy components (DataGrid, ResultDetailPanel, SearchManager panel, individual layouts) are lazy-loaded via `React.lazy`. First click on a non-default layout takes 100–300ms longer; subsequent clicks are instant.
3. Slow result render in DataGrid → check `selectedPropertiesCollection`. Too many columns × too many rows × heavy cell renderers (Persona, Taxonomy) compound.
4. AbortController is wired — typing fast shouldn't queue duplicate requests. If you see queued requests in Network tab, the orchestrator isn't aborting (file a bug).

### Resolution

| Cause | Fix |
|---|---|
| Initial bundle too large | Check `release/analysis-logs/bundle-sizes.json` after build; if a web part broke its budget, file a follow-up |
| Slow layout switch | Pre-hover the layout button to trigger preload (`onMouseEnter` triggers lazy import). T1.D11. |
| DataGrid scrolling janky on mobile | iOS momentum scrolling enabled in current build; if jerky, hard-refresh + clear cache |
| Refiner panel flickers during fast typing | Refiner stability mode debounces `displayRefiners` — old bundle if you still see flicker |

## Diagnostic tooling

### `?debug=1` DebugFab

Append `?debug=1` to any page hosting SP Search. The **floating action button (FAB)** in the bottom-right opens a panel with:

- **Store** — live state for every slice (query, filters, results, UI)
- **Network** — every search/refiner/vertical request with timing
- **Multi-Context** — every context on the page (refcount, init status, URL sync, registered web parts)
- **Init order** — which web part registered first, whether Filters arrived after Results' first search
- **Force Dispose** — manually dispose a context (recovery from a stuck state)

Only Owners/Admins see the FAB. End users with `?debug=1` see nothing.

### Browser console

All SP Search log lines are prefixed `[SP Search]`. Filter the console by `[SP Search]` to isolate. Warn/error level always emits; debug/info is gated to `?debug=1` in production builds.

### PowerShell

```powershell
# Check site has all hidden lists
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/search" -Interactive
@('SearchSavedQueries','SearchHistory','SearchCollections') | ForEach-Object {
    Get-PnPList -Identity $_ -ErrorAction SilentlyContinue | Select Title, ItemCount, EnableVersioning
}

# Verify item-level security on SearchHistory
$historyList = Get-PnPList -Identity "SearchHistory"
"$($historyList.Title): ReadSecurity=$($historyList.ReadSecurity), WriteSecurity=$($historyList.WriteSecurity)"
# Expect: ReadSecurity=2, WriteSecurity=2

# Verify SearchHistory runtime fields are present
@(
  'QueryText',
  'QueryHash',
  'Vertical',
  'SearchPageUrl',
  'SearchState',
  'UseCount',
  'ResultCount',
  'IsZeroResult',
  'ClickedItems',
  'SearchTimestamp'
) | ForEach-Object {
  Get-PnPField -List SearchHistory -Identity $_ -ErrorAction Stop | Select InternalName, TypeAsString, Indexed
}

# Find managed properties available for refinement
Invoke-PnPSearchQuery -Query "*" -SelectProperties "Title" -SortList @{} -TrimDuplicates $false |
    Select -ExpandProperty PrimaryQueryResult |
    Select -ExpandProperty RefinementResults
```

### SharePoint admin centre paths to know

| Surface | Path | Use for |
|---|---|---|
| API access | SP admin centre → Advanced → **API access** | Approving Graph permissions (`People.Read`, `User.Read`, `User.Read.All`) |
| Search schema | Site Settings → **Search Schema** (or tenant-level via SP admin centre → Search → Manage Search Schema) | Marking managed properties as refinable / sortable / queryable |
| Content sources | SP admin centre → Search → **Manage Content Sources** | Verifying the target site is being crawled |
| Crawl log | SP admin centre → Search → **Crawl log** | Diagnosing why specific items aren't indexed |
| Query rules | SP admin centre → Search → **Manage Query Rules** | Promoted results / best bets |

## Escalation

If a symptom doesn't match anything above and the Pre-Flight passes:

1. Capture `?debug=1` → Multi-Context tab screenshot (shows store state)
2. Capture browser console errors filtered to `[SP Search]`
3. Capture the failing request from Network tab (URL + response headers + body)
4. File a bug at the GitHub issue tracker with all three attached.
