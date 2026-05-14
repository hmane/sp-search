# SP Search — Multi-Context Authoring Guide (T3.D5)

> Owned by T3 / Multi-Context Mastery. Walkthrough for putting
> two (or more) independent search experiences on a single
> SharePoint page.

## What is a search context?

Every SP Search web part declares a **`searchContextId`** property
(page 1, group 1, first field). Web parts that share an ID also
share one Zustand store, one orchestrator, one URL-sync subscription,
and one set of registered providers / actions / layouts.

Web parts with different IDs each get their own store. They never
exchange queries, filters, or vertical selections — even on the
same page.

The default ID is the literal string `"default"`. Single-search
pages don't need to change anything. Multi-context pages set
distinct IDs (e.g. `"hr-search"` + `"policy-search"`) so the two
experiences run side-by-side without interference.

## When to use multi-context

Three concrete shapes are worth knowing:

1. **Side-by-side scopes** — One context indexes the entire tenant
   (Knowledge Base preset); the other narrows to the current site
   (Documents preset). The
   `test-multi-context-tenant-vs-site` provisioning sample
   demonstrates this exact shape — see
   [provisioning-guide.md](./provisioning-guide.md).

2. **Different verticals on different page regions** — A page with
   a People search in the header and a Documents search in the
   body. Each has its own filter sidebar.

3. **Different data providers per context** — Context A routes to
   SharePoint Search; Context B routes to Graph Search. The
   `dataProviderId` per-vertical override (T3.D7) handles the
   per-row case; the searchContextId fork handles the per-region
   case.

## Authoring checklist

For every web part you place on the multi-context page:

1. Open the property pane → first group ("Search context").
2. Set **Search Context ID** to the same value across web parts you
   want connected (e.g. all Context A web parts use `"docs-site"`,
   all Context B web parts use `"kb-tenant"`).
3. If you mistype or leave an ID stale, an edit-mode MessageBar on
   the affected web part shows the mismatch (T3.D2) — the banner
   lists the IDs in play on the page so you can spot the typo.
4. Verify the URL after typing in one of the boxes: each context
   gets its own URL prefix so deep links round-trip cleanly
   (T3.D3 + T3.D6).

## URL parameter namespacing

With one context per page the URL is clean:

```
https://contoso.sharepoint.com/sites/search/SitePages/Search.aspx?q=annual+report&v=documents&p=2
```

With **two or more** contexts, each context's parameters are prefixed
with a short stable namespace derived from the searchContextId:

```
https://contoso.sharepoint.com/sites/search/SitePages/Search.aspx?docs-site.q=annual+report&kb-tenant.q=policy
```

Admins who want shorter, hand-typeable prefixes can override the
auto-computed namespace via the runtime option
`initializeSearchContext(id, ctx, { urlPrefix: 'ctx1' })` (T3.D6).
Property pane wiring is a v1.1 follow-up; today the override is
available to third-party extensions that bootstrap their own
contexts.

To opt a context out of URL sync entirely (e.g. an embedded
"saved-search runner" widget that must not stomp the page URL),
pass `{ enableUrlSync: false }` to the same call (T3.D6).

## URL alias uniqueness on filter values

When two managed properties on the same filterConfig normalize to
the same short alias (e.g. `Author` and `AuthorOWSUSER` both
shortening to `au`), the URL-alias collision validator (T3.D3)
disambiguates them at config-load time. A deep link round-trips
cleanly even with both filters enabled.

If an admin configures two filters that share a stem, the second
filter's alias bumps by a digit (`au1`, `au2`, ...) — the
mapping is deterministic so URLs stay stable across page loads.

## Edit-mode diagnostics

Three edit-mode warnings surface multi-context misconfiguration:

1. **Mismatch banner (T3.D2)** — When the IDs configured across
   web parts on the page don't agree (e.g. Box has `default` while
   Results has `hr-search`), every affected web part renders a
   MessageBar listing the IDs in play. The banner only fires in
   edit mode; view mode hides it. Fix by harmonising the
   property-pane values.

2. **Init-order warning (T3.D10)** — When Results' first search
   runs before the Filters web part has loaded (URL-deep-linked
   filter values silently failed to apply), an edit-mode
   MessageBar offers a Retry button. Clicking Retry re-fires the
   search now that filterConfig is populated.

3. **DataProvider id validation (T3.D7)** — When a vertical
   references a `dataProviderId` not in the registered provider
   list, an edit-mode MessageBar names the bad vertical + offers
   a Did-You-Mean against the registered ids (`sharepoint-search`,
   `graph-search`, `graph-people`).

## Inspecting context state at runtime

Activate the cross-bundle DebugFab on any page via
`?debug=1`. The Debug Panel has a **Multi-Context** tab (T3.D8)
that enumerates every context on the page with:

- Context ID
- URL prefix (auto-computed + admin override marker when present)
- Refcount (number of web parts holding the context alive)
- Init flag (orchestrator started)
- URL sync attachment status
- Registered web parts (which web parts wrote this context's ID)
- Live store snapshot (queryText / vertical / activeFilters count /
  items.length)
- "Force dispose" button per row (bypasses refcount — for
  debugging stuck contexts only)

The tab auto-refreshes every 2 seconds so multi-context bleed
scenarios are visible in real time.

## Lifecycle

Each web part is a refcount holder for its context (T3.D1). When
the last web part on a context unmounts (page navigation in the
SharePoint Modern SPA shell), the context is automatically
disposed after a microtask defer (the defer lets a new mount with
the same ID re-claim the context before teardown — handles the
SPA navigation race where the next page's onInit fires before the
prior page's onDispose).

`window.__sp_search_context_map__.size` returns to zero after the
last unmount on a page navigation. Verify in DevTools console if
you suspect a leak.

Third-party extensions that import `getStore` directly must follow
the same refcount contract — see
[extensibility-guide.md](./extensibility-guide.md#lifecycle-refcounted-context-dispose-t3d1).

## Common failure modes + recovery

| Symptom | Cause | Fix |
|---------|-------|-----|
| Two web parts on the same page don't talk to each other | searchContextId mismatch | Open each web part's property pane → first group → harmonise the IDs. Edit-mode MessageBar lists the IDs in play. |
| URL deep link doesn't apply filters on first load | Results loaded before Filters | Click Retry on the init-order MessageBar (T3.D10) OR reload the page. |
| Two managed properties' URL deep links collide silently | Aliases share a stem | T3.D3 validation flags this in edit mode. Rename one alias or the validator auto-disambiguates with a digit suffix. |
| People vertical returns documents instead of profile cards | `dataProviderId` typo | T3.D7 validator flags Did-You-Mean. Verticals property pane → Manage refiners → dataProviderId column → pick from the dropdown (`graph-people` for the People vertical). |
| Context state survives across page navigation | refcount not released | T3.D1 should handle this. Verify in DebugPanel Multi-Context tab; "Force dispose" if stuck. |

## Sample pages

Two provisioning samples demonstrate the moat:

- `test-multi-context` — Documents context + People context on
  one page, same scope.
- `test-multi-context-tenant-vs-site` — Tenant-wide Knowledge
  Base + site-scoped Documents on one page, different scopes.

Both ship via `scripts/Provision-TestPages.ps1`. See
[provisioning-guide.md](./provisioning-guide.md).
