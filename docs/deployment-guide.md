# SP Search — Deployment Guide

Step-by-step instructions for building, deploying, and configuring SP Search in a SharePoint Online environment.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **SharePoint Online** | Microsoft 365 tenant with SharePoint Online |
| **Node.js** | v18.x (LTS) — required by SPFx 1.21.1 |
| **Gulp CLI** | `npm install -g gulp-cli` |
| **PnP.PowerShell** | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| **Permissions** | Site Collection Admin on the target site (for app deployment + list provisioning) |
| **App Catalog** | Site-level or tenant-level App Catalog enabled |

---

## Step 1: Build the Solution

```bash
# Navigate to the project root
cd /path/to/sp-search

# Install dependencies (if not already done)
npm install

# Build the spfx-toolkit dependency (if using local link)
cd /path/to/spfx-toolkit && npm run build && cd /path/to/sp-search

# Production build
gulp bundle --ship

# Package the solution
gulp package-solution --ship
```

The `.sppkg` file is generated at:
```
sharepoint/solution/sp-search.sppkg
```

### Verify Build

- `gulp bundle --ship` should complete with **0 errors**
- Check `dist/` for expected chunks (entry bundles + lazy-loaded chunks + vendor chunks)
- The `.sppkg` file should be present in `sharepoint/solution/`

---

## Step 2: Deploy to App Catalog

### Option A: Tenant-Level App Catalog

1. Navigate to your tenant App Catalog site: `https://<tenant>.sharepoint.com/sites/appcatalog`
2. Go to **Apps for SharePoint**
3. Upload `sp-search.sppkg`
4. In the trust dialog, check **Make this solution available to all sites in the organization** if you want tenant-wide availability
5. Click **Deploy**

### Option B: Site-Level App Catalog

1. Ensure the target site has a site-level App Catalog enabled:
   ```powershell
   Add-PnPSiteCollectionAppCatalog -Site "https://<tenant>.sharepoint.com/sites/<target>"
   ```
2. Navigate to: `https://<tenant>.sharepoint.com/sites/<target>/AppCatalog`
3. Upload `sp-search.sppkg`
4. Click **Deploy**

### Add the App to the Site

1. Navigate to the target site
2. Go to **Site Contents** > **New** > **App**
3. Find **SP Search** and click **Add**
4. Wait for installation to complete

---

## Step 3: Provision Hidden Lists

The 4 hidden lists required by SP Search must be created via PowerShell before using saved searches, history, collections, or promoted results.

```powershell
# Connect to the target site
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<target>" -Interactive

# Run the provisioning script
.\scripts\Provision-SPSearchLists.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<target>"
```

See [provisioning-guide.md](./provisioning-guide.md) for full parameter documentation and troubleshooting.

### What Gets Created

| List | Purpose |
|------|---------|
| **SearchSavedQueries** | Saved and shared search queries |
| **SearchHistory** | Per-user search history (auto-logged, auto-pruned) |
| **SearchCollections** | Pinboard collections of search results |
| **SearchConfiguration** | Admin config: scopes, promoted results, state snapshots |

### Post-Provisioning Steps

1. **SearchConfiguration admin access**: Add your admin security group as list owners for write access. The script prints a reminder with instructions.
2. **Verify indexes**: Confirm that SearchHistory has indexes on Author and SearchTimestamp (critical for list view threshold performance).

---

## Step 4: Configure Web Parts

### Minimal Setup (Search Box + Results)

1. Edit a SharePoint page
2. Add the **SP Search Box** web part
3. Set `searchContextId` = `"default"`
4. Add the **SP Search Results** web part below it
5. Set `searchContextId` = `"default"` (must match)
6. Configure `selectedProperties` with the properties you need
7. Publish the page

### Full Setup (All 5 Web Parts)

Add all web parts to the page and set the same `searchContextId` on each:

| Web Part | Placement | Key Config |
|----------|-----------|------------|
| **Search Verticals** | Top of search area | Configure `verticals` JSON |
| **Search Box** | Below verticals | Enable features as needed |
| **Search Filters** | Left sidebar | Set `applyMode` |
| **Search Results** | Main content area | Set `queryTemplate`, `selectedProperties`, `pageSize` |
| **Search Manager** | Right sidebar or panel mode | Set `mode` |

See [admin-guide.md](./admin-guide.md) for detailed property documentation.

---

## Step 5: Verify Deployment

### Smoke Tests

| Test | Expected Result |
|------|-----------------|
| Type a query in Search Box | Results appear in Search Results web part |
| Click a filter value | Results refine; filter pill appears above results |
| Click a vertical tab | Results change; badge counts update |
| Click a result | Detail panel opens with preview |
| Switch layout (List / Compact / Grid) | Results re-render in selected layout |
| Press browser Back button | Previous search state restores |
| Copy URL and open in new tab | Same search state loads |

### Feature Tests (Requires Provisioned Lists)

| Test | Expected Result |
|------|-----------------|
| Save a search | Appears in Search Manager > Saved tab |
| Check search history | Recent searches listed in Search Manager > History |
| Pin a result to collection | Appears in Search Manager > Collections tab |
| Share a search to another user | Recipient sees it in Shared tab |

### Performance Checks

| Metric | Target |
|--------|--------|
| Initial page load | < 3 seconds (warm cache) |
| Query-to-render | < 1 second |
| DataGrid with 500+ results | Smooth virtual scrolling |
| No memory leaks | Stable memory on repeated searches |

---

## Bundle Size Reference

Ship build (minified) — February 2026:

### Entry Bundles (Per Web Part)

| Bundle | Size | Notes |
|--------|------|-------|
| sp-search-store-library | 14 KB | Shared store + registries |
| sp-search-verticals-web-part | 775 KB | Lightest web part |
| sp-search-box-web-part | 1.0 MB | Includes query builder |
| sp-search-results-web-part | 1.1 MB | Base + List/Compact layouts |
| sp-search-filters-web-part | 1.4 MB | Base + checkbox/toggle filters |
| sp-search-manager-web-part | 1.4 MB | Base + CRUD services |

### Lazy-Loaded Chunks (On Demand)

| Chunk | Size | Loaded When |
|-------|------|-------------|
| CardLayout | 122 KB | User selects Card layout |
| DevExtremeDataGrid | 71 KB | User selects DataGrid layout |
| SearchManager | 46 KB | Search Manager panel opens |
| TaxonomyTreeFilter | 50 KB | Page has taxonomy filter configured |
| VisualFilterBuilder | 45 KB | Admin enables visual filter builder |
| ResultDetailPanel | 27 KB | User clicks a result |
| PeopleLayout | 25 KB | User selects People layout |
| PeoplePickerFilter | 17 KB | Page has people picker filter |
| TagBoxFilter | 12 KB | Page has tag box filter |
| GalleryLayout | 3 KB | User selects Gallery layout |
| DataGridLayout (wrapper) | 2.1 KB | User selects DataGrid layout |

### Vendor Chunks (Shared, Cached)

| Library | Chunks | Total Size | Notes |
|---------|--------|------------|-------|
| DevExtreme | 8 chunks | ~9.4 MB | Loaded only when DataGrid/TreeView/FilterBuilder used |
| Fluent UI | 4 chunks | ~11.4 MB | Shared across SPFx; often already cached by SharePoint |
| PnP Controls | 1 chunk | 3.3 MB | PeoplePicker + TaxonomyPicker |
| spfx-toolkit | 2 chunks | 273 KB | Persona + Tooltip utilities |
| React Hook Form | 1 chunk | 74 KB | Form handling |

**Note:** DevExtreme and Fluent UI vendor chunks appear large but are heavily cached by the browser. On repeat visits, only the entry bundle and app-specific chunks are downloaded. SharePoint pages typically already have Fluent UI loaded, so those chunks are often cache hits.

### Locale Files

PnP Controls ships ~20 locale files at ~24 KB each (~480 KB total). Only the user's locale is loaded at runtime.

---

## Troubleshooting

### Build Errors

| Error | Solution |
|-------|----------|
| `Cannot find module 'spfx-toolkit/lib/...'` | Rebuild spfx-toolkit: `cd /path/to/spfx-toolkit && npm run build` |
| TypeScript errors in `.d.ts` files | Run `npm install` to ensure type definitions are current |
| `gulp bundle` timeout | Increase Node.js memory: `export NODE_OPTIONS=--max-old-space-size=8192` |

### Runtime Errors

| Error | Solution |
|-------|----------|
| "No search data provider registered" | Ensure `searchContextId` matches across web parts; Search Box registers the provider |
| Filters show no values | Verify managed properties are set as **Refinable** in Search Schema |
| Empty search results | Check `queryTemplate` syntax; verify `selectedProperties` are **Retrievable** |
| "SearchManagerService is not ready" | `currentUser()` call failed — check permissions; service retries on next `initialize()` call |
| History/saved searches not working | Verify hidden lists are provisioned and user has Add Items permission |

### List Threshold Issues

The SearchHistory list will exceed 5,000 items on active sites. All queries are designed to filter by `Author eq [Me]` as the first CAML predicate, which uses the Author index. If you see threshold errors:

1. Verify the Author index exists on SearchHistory
2. Run the provisioning script again (idempotent) to recreate missing indexes
3. Configure history cleanup TTL in SearchConfiguration (30/60/90 days)
