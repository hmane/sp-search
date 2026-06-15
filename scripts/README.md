# SP Search — Provisioning Scripts

Production install + admin tooling for SP Search. Run from this directory unless noted.

## Which script do I run?

**To install SP Search on a brand-new site → [`Deploy-SPSearchSolution.ps1`](Deploy-SPSearchSolution.ps1).** That's the canonical one-shot install. It uploads the `.sppkg` to the target site's app catalog, installs the app, applies the declarative PnP provisioning template (`provisioning/SiteTemplate.xml`), and configures item-level security on `SearchHistory`.

```powershell
.\scripts\Deploy-SPSearchSolution.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<your-aad-app-id>" `
    -ProvisionSite
```

If `-ProvisionSite` fails (the PnP provisioning engine occasionally rejects valid templates after a schema version bump), fall back to the imperative path in [Fallback install](#fallback-install) below.

After the install, optionally:

- **Map crawled → managed properties** via `Map-CrawledProperties.ps1` — required once per tenant if your search schema differs from the defaults.
- **Wire a scenario preset** via `Search-ScenarioPresets.ps1` — applies a pre-configured layout + filter set (Documents / News / People / Knowledge Base / Hub Search / Policy Search) to a page.
- **Promote page configuration** via `Export-SPSearchPageConfig.ps1` and `Import-SPSearchPageConfig.ps1` — exports/imports all SP Search web part property bags as JSON for environment moves.

## Script inventory

| Script | When to run |
|---|---|
| `Deploy-SPSearchSolution.ps1` | Canonical install. Run once per target site. |
| `Map-CrawledProperties.ps1` | Tenant-level schema configuration; idempotent. |
| `Search-ScenarioPresets.ps1` | Optional: applies a scenario preset to an existing page. |
| `Export-SPSearchPageConfig.ps1` | Exports raw SP Search web part properties from a modern page to JSON. |
| `Import-SPSearchPageConfig.ps1` | Imports a saved JSON config into matching SP Search web parts on a target page. |
| `Setup-SPSearchSite.ps1` | **Fallback** imperative install (pre-template). See below. |
| `Provision-SPSearchLists.ps1` | **Fallback** lists-only provisioning. Wrapped by `Setup-SPSearchSite.ps1`. |
| `Provision-SPSearchPage.ps1` | **Fallback** page-only provisioning. Wrapped by `Setup-SPSearchSite.ps1`. |
| `check-bundle-sizes.js` | CI gate. Run by `npm run check:bundle-sizes` after `npm run package`. Not for tenants. |

## Promote page configuration between environments

Use this when a configured search page needs to move from dev to test or prod without manually re-entering every web part property.

Export from the source page:

```powershell
.\scripts\Export-SPSearchPageConfig.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search-dev" `
    -ClientId "<your-aad-app-id>" `
    -PageName "Search" `
    -OutputPath ".\config\search-page.dev.json" `
    -TokenizeSiteUrl `
    -Force
```

Import into the target page:

```powershell
.\scripts\Import-SPSearchPageConfig.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search-prod" `
    -ClientId "<your-aad-app-id>" `
    -PageName "Search" `
    -ConfigPath ".\config\search-page.dev.json" `
    -WhatIf

.\scripts\Import-SPSearchPageConfig.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search-prod" `
    -ClientId "<your-aad-app-id>" `
    -PageName "Search" `
    -ConfigPath ".\config\search-page.dev.json" `
    -Force
```

The import updates existing SP Search web parts only. Provision the target page first with `Deploy-SPSearchSolution.ps1 -ProvisionSite`, `Provision-SPSearchPage.ps1`, or `Search-ScenarioPresets.ps1`, then import the saved JSON. The JSON covers the complete SPFx `properties` object for Box, Results, Filters, Results + Filters, Verticals, Search Manager, and Admin Manager, including collection fields such as refiners, verticals, grid columns, selected properties, coverage profiles, and audience targeting. It does not migrate hidden-list data, saved searches, history, collections, page sections, or non-SP Search web parts.

Environment-specific values can be tokenized. Export `-TokenizeSiteUrl` replaces the source site URL with `{siteUrl}`; import always replaces `{siteUrl}` with the target `-SiteUrl`. Additional tokens can be supplied with `-TokenFile`.

See [`config/sp-search-page-config.sample.json`](../config/sp-search-page-config.sample.json) for a representative full export with the standard six-web-part page, collection fields, fake instance IDs, and tokenized URLs. Use `Provision-SPSearchPage.ps1 -UseFullWidthExperience` when you want the optional combined Results + Filters wrapper instead of separate Results and Filters web parts.

## Fallback install

If the declarative template path errors out, run `Setup-SPSearchSite.ps1` instead — it chains the three imperative `Provision-*.ps1` scripts and produces equivalent state without going through `Invoke-PnPSiteTemplate`. Both paths land on the same hidden lists, the same Search.aspx page, and the same web part GUIDs.

```powershell
.\scripts\Setup-SPSearchSite.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/search" `
    -ClientId "<your-aad-app-id>"
```

## Prerequisites

- **PnP.PowerShell 3.x** (`Install-Module PnP.PowerShell`).
- **Azure AD app registration** with `Sites.FullControl.All` granted, OR interactive sign-in with a SharePoint site administrator account.
- **`.sppkg` built** via `npm run package`. The script defaults to `sharepoint/solution/sp-search.sppkg` relative to repo root.

See [`docs/deployment-guide.md`](../docs/deployment-guide.md) for the full deployment runbook and [`docs/provisioning-guide.md`](../docs/provisioning-guide.md) for what the PnP template creates.
