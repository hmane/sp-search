# SP Search — Deployment Guide

This guide covers build, package, deployment, provisioning, permissions, and first-page setup for the current solution.

## Prerequisites

| Requirement | Details |
|-------------|---------|
| Node.js | `>=22.14.0 <23.0.0` |
| PnP.PowerShell | `Install-Module PnP.PowerShell -Scope CurrentUser` |
| SharePoint permissions | Site Collection Admin on the target site |
| App Catalog | Tenant-level or site-level App Catalog |

## Build and Package

```bash
npm install
npm run type-check
npm test
npm run package
```

Package output:

```text
sharepoint/solution/sp-search.sppkg
```

## Deploy the Solution

### Tenant App Catalog

1. Upload `sharepoint/solution/sp-search.sppkg` to the tenant App Catalog.
2. Deploy the solution.
3. Optionally make it available to all sites in the organization.

### Site App Catalog

1. Enable a site-level App Catalog if needed.
2. Upload `sp-search.sppkg`.
3. Deploy the solution to that site.

### Add the App

After deployment, add the SP Search app to the target site from **Site contents**.

### Automated deploy (T4.D10)

`scripts/Deploy-SPSearchSolution.ps1` automates upload + deploy
against either a site- or tenant-level App Catalog:

```powershell
# Site-level (default) — uploads to <SiteUrl>/AppCatalog
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>"

# Tenant-level — requires -AppCatalogUrl
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>" `
    -AppCatalogScope TenantLevel `
    -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog"

# Deploy from a published release artifact (Azure DevOps, GitHub,
# any direct https URL) instead of a locally-built .sppkg
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>" `
    -AppCatalogScope TenantLevel `
    -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog" `
    -ReleaseArtifactUrl "https://dev.azure.com/.../sp-search-v1.0.0.sppkg"
```

`-ReleaseArtifactUrl` downloads the `.sppkg` to a temp file and
overrides the default `-PackagePath`. Useful when an admin wants to
deploy the canonical release artifact without cloning the repo and
rebuilding.

## Provision Hidden Lists

The Search Manager, history, collections, saved searches, health, and insights features depend on the hidden lists created by the provisioning script.

```powershell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/search" -Interactive
.\scripts\Provision-SPSearchLists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/search"
```

The script provisions:

- `SearchSavedQueries`
- `SearchHistory`
- `SearchCollections`

See [provisioning-guide.md](./provisioning-guide.md) for schema details.

## Safe Re-Runs (T4.D1)

The four provisioning scripts that touch destructive paths now implement the
standard PowerShell `SupportsShouldProcess` contract:

| Script | Destructive paths | Default behaviour |
|--------|-------------------|-------------------|
| `Setup-SPSearchSite.ps1` | `Remove-PnPPage` (existing search page); `Remove-PnPField` when a UserMulti column was previously created as the wrong type | Prompts before each destructive op |
| `Provision-SPSearchLists.ps1` | `Set-PnPList -BreakRoleInheritance` on each of the three hidden lists | Prompts before each permission reset |
| `Map-CrawledProperties.ps1` | `Set-PnPSearchConfiguration -Scope Site` (site-scoped search schema overwrite) | Prompts before each mapping |
| `Provision-TestData.ps1` | `Remove-PnPList` of the seeded libraries/lists when `-CleanExisting` is set | Prompts before the batch removal |

Common patterns:

- **Preview first** — add `-WhatIf` to any of the four scripts to see the
  destructive operations the script would perform without executing them. The
  output lines are tagged `What if: Performing the operation 'X' on target
  'Y'.`
- **Bypass the prompt** — add `-Force` for CI / scripted callers. `-Force`
  short-circuits the `ShouldProcess` gate; the destructive op still runs.
- **Decline interactively** — leave both off. PowerShell asks `Y/N` before
  each destructive op; the script prints a "preserved / re-run with `-Force`"
  message and continues (`Setup-SPSearchSite.ps1` exits non-zero if the page
  removal is declined, because the script cannot continue without removing
  the existing page).

PSScriptAnalyzer's `PSShouldProcess` rule reports zero violations across the
four scripts.

## Optional: Provision a Scenario Search Page

The repo also includes `scripts/Search-ScenarioPresets.ps1`, which can provision complete search pages for built-in scenarios.

Useful commands:

```powershell
Get-SearchScenarioPresetList
Get-SearchScenarioPreset -Name "documents"
Invoke-SearchScenarioPage -SiteUrl "https://contoso.sharepoint.com/sites/search" -Name "knowledge-base" -PageName "knowledge-search.aspx"
```

Built-in scenario page presets:

- `general`
- `documents`
- `people`
- `news`
- `media`
- `hub-search`
- `knowledge-base`
- `policy-search`

## Graph Permissions

Grant Graph permissions before enabling Graph-backed People features in production.

| Capability | Permission |
|------------|------------|
| Graph people vertical | Microsoft Graph access for `/search/query` on people entities |
| Org chart section | `User.Read.All` |

If Graph permission is not granted:

- SharePoint-backed search still works
- Graph people verticals do not return full Graph people results
- org-chart relationships stay hidden

## Minimal Page Setup

The starter manifests now ship useful defaults, so a minimal page no longer needs immediate JSON editing.

1. Add **SP Search Box**
2. Add **SP Search Results**
3. Set the same `searchContextId` on both web parts
4. Publish

Starter behavior:

- site-scoped query
- query template `{searchTerms}`
- List, Compact, and Grid layouts
- 10 results per page with paging enabled
- result count and sort enabled

## Full Search Page Setup

Recommended authoring order:

1. **SP Search Verticals**
2. **SP Search Box**
3. **SP Search Filters**
4. **SP Search Results**
5. **SP Search Manager** in panel mode

Then set the same `searchContextId` on all five web parts.

Recommended starter page behavior:

- Results preset: `general` or `documents`
- Filters starter set: file type, modified date, author
- Verticals starter set: all, documents, pages, sites
- Manager mode: `panel`

## Smoke Test Checklist

Run this after deployment:

- After uploading the `.sppkg`, the SharePoint admin center "API access" page surfaces a pending `People.Read` request for Microsoft Graph (declared in `webApiPermissionRequests` per Found.D10). Approve it before deploying the People vertical (Graph-backed search results).

| Test | Expected result |
|------|-----------------|
| Type a query | Results load and no false empty state flashes |
| Switch verticals | Tabs update query and counts |
| Apply author filter | People picker refines results correctly |
| Switch to Grid | Dynamic columns, chooser, resize, export, fullscreen all work |
| Export CSV/XLSX | Download contains visible grid rows or selected rows |
| Open Health tab | Zero-result queries load if history exists |
| Open Insights tab | Trend cards and charts load |
| Open a People result | Graph people card actions work, org chart expands if permission exists |

## Troubleshooting

| Issue | Action |
|-------|--------|
| No provider registered | Check `searchContextId` consistency and ensure Results/Box are on the page |
| Filters show no values | Confirm managed properties are refinable and included in filter config |
| Author people filter returns nothing | Use `AuthorOWSUSER`, not `Author` |
| Graph people results missing | Verify Graph permission approval and per-vertical `dataProviderId` |
| History/Health/Insights empty | Ensure provisioning script completed and hidden lists exist |
| DevExtreme font or CSS issues in local serve | Verify the `fast-serve/webpack.extend.js` overrides, then restart `npm run start` |

## Related Docs

- [admin-guide.md](./admin-guide.md)
- [provisioning-guide.md](./provisioning-guide.md)
- [pnp-modern-search-alignment.md](./pnp-modern-search-alignment.md)
- [product-cleanup-audit.md](./product-cleanup-audit.md)
