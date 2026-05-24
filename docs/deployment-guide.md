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

**Recommended path — one script does everything:**

`scripts/Deploy-SPSearchSolution.ps1` uploads the `.sppkg`, installs the app on the target site, applies the PnP provisioning template (creates the three hidden lists + the Search.aspx page with all web parts pre-wired), and configures `SearchHistory` item-level security. Use it against either a site- or tenant-level App Catalog. See [`scripts/README.md`](../scripts/README.md) for the full script inventory.

```powershell
# Site-level (default) — uploads to <SiteUrl>/AppCatalog
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>" `
    -ProvisionSite

# Tenant-level — requires -AppCatalogUrl
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>" `
    -AppCatalogScope TenantLevel `
    -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog" `
    -ProvisionSite

# Deploy from a published release artifact (Azure DevOps, GitHub,
# any direct https URL) instead of a locally-built .sppkg
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>" `
    -AppCatalogScope TenantLevel `
    -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog" `
    -ReleaseArtifactUrl "https://dev.azure.com/.../sp-search-v1.0.0.sppkg" `
    -ProvisionSite
```

`-ReleaseArtifactUrl` downloads the `.sppkg` to a temp file and overrides the default `-PackagePath`. Drop it for a local-build install.

`-ProvisionSite` is what wires the lists + page in one go. Omit it if you want to install the app and hand-build pages instead.

### Manual install (when you can't run the script)

If your org requires the `.sppkg` to be uploaded manually (governance, audit trail, etc.):

**Tenant App Catalog path:**

1. Upload `sharepoint/solution/sp-search.sppkg` to the tenant App Catalog.
2. Deploy the solution; optionally make it available to all sites in the organization.
3. On the target site, **Site contents → New → App → SP Search**.
4. Run the imperative provisioning fallback (see below).

**Site App Catalog path:**

1. Enable a site-level App Catalog if needed.
2. Upload `sp-search.sppkg`; deploy the solution to that site.
3. **Site contents → New → App → SP Search**.
4. Run the imperative provisioning fallback (see below).

### Imperative provisioning fallback

If you can't run `Deploy-SPSearchSolution.ps1 -ProvisionSite` (or it fails on the `Invoke-PnPSiteTemplate` step), run the imperative wrapper instead. It chains lists + page creation without going through the PnP provisioning engine.

```powershell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/search" -Interactive
.\scripts\Setup-SPSearchSite.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<azure-ad-app-id>"
```

`Setup-SPSearchSite.ps1` provisions the three hidden lists (`SearchSavedQueries`, `SearchHistory`, `SearchCollections`), creates the Search.aspx page, and wires the five web parts. See [provisioning-guide.md](./provisioning-guide.md) for schema details.

## Safe Re-Runs (T4.D1)

The four provisioning scripts that touch destructive paths now implement the
standard PowerShell `SupportsShouldProcess` contract:

| Script | Destructive paths | Default behaviour |
|--------|-------------------|-------------------|
| `Deploy-SPSearchSolution.ps1` | `Invoke-PnPSiteTemplate` (overwrites Search.aspx + the three hidden lists if `Overwrite="true"` in the template); `Set-PnPList` updates on `SearchHistory` item-level security | Prompts before each destructive op |
| `Setup-SPSearchSite.ps1` (fallback) | `Remove-PnPPage` (existing search page); `Remove-PnPField` when a UserMulti column was previously created as the wrong type | Prompts before each destructive op |
| `Provision-SPSearchLists.ps1` (fallback) | `Set-PnPList -BreakRoleInheritance` on each of the three hidden lists | Prompts before each permission reset |
| `Map-CrawledProperties.ps1` | `Set-PnPSearchConfiguration -Scope Site` (site-scoped search schema overwrite) | Prompts before each mapping |

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
| Export CSV/XLSX | Download contains visible grid rows |
| Click a PDF or Office result | In-page preview Modal opens (clickTarget=`panel`); document renders without breaking out of the page |
| Open detail panel (right-arrow icon) | Side panel slides in with preview + metadata + Previous/Next arrows |
| Open Admin Manager → Pre-Flight | Tenant readiness checklist runs; surfaces any failed checks (Graph permission, hidden lists, schema mappings) |
| Open Admin Manager → Dashboard | Content Coverage, Search Quality, and Zero-Result Queries sections render (after some search activity has accumulated) |
| Open Health tab | Zero-result queries load if history exists |
| Open Insights tab | Trend cards and charts load |
| Open a People result | Graph people card actions work, org chart expands if permission exists |

## Troubleshooting

Quick reference. For symptom→diagnosis→resolution depth, see [admin-runbook.md](./admin-runbook.md).

| Issue | First check |
|-------|------------|
| No provider registered | `searchContextId` consistency across Box/Results/Filters/Verticals/Manager |
| Filters show no values | Managed property is marked **refinable** in the SharePoint search schema |
| Author people filter returns nothing | Use `AuthorOWSUSER`, not `Author` |
| Graph people results missing | Graph permission approved in SP admin centre → Advanced → API access |
| History / Health / Insights empty | Hidden lists provisioned (re-run `Deploy-SPSearchSolution.ps1 -ProvisionSite` or `Setup-SPSearchSite.ps1`) |
| Web part doesn't load on the page | `?debug=1` to open the DebugFab + check browser console; check `searchContextId` mismatch banner in edit mode |
| Click a PDF, current tab navigates away | Page was published before `data-interception="off"` shipped; re-publish the page |
| "This page has been blocked by Chrome" inside modal preview | Old bundle cached; hard-refresh (Cmd+Shift+R) — `<embed>` for PDFs ships in the current build |
| DevExtreme font or CSS issues in local serve | Verify the `config/spfx-customize-webpack.js` overrides + `gulpfile.js` rule patches, then restart `npm run start` |

## Related Docs

- [admin-guide.md](./admin-guide.md)
- [admin-runbook.md](./admin-runbook.md) — full symptom→resolution playbook
- [end-user-guide.md](./end-user-guide.md)
- [provisioning-guide.md](./provisioning-guide.md)
- [scripts/README.md](../scripts/README.md) — script inventory
- [pnp-modern-search-alignment.md](./pnp-modern-search-alignment.md)
