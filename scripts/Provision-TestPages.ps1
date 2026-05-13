<#
.SYNOPSIS
    One-stop test page provisioning for SP Search — creates multiple pages
    covering all search scenarios, filter types, layouts, and edge cases.

.DESCRIPTION
    Provisions a complete test suite of SharePoint pages, each exercising a
    specific combination of SP Search web part capabilities. Run this after
    deploying the .sppkg and provisioning test data / search lists.

    Pages created:
      1. test-general          General search (all content, list layout)
      2. test-documents        Document search (grid + compact, file-type refiners)
      3. test-people           People search (Graph provider, people layout)
      4. test-news             News search (card layout, date-sorted)
      5. test-media            Media gallery (gallery layout, image/video)
      6. test-multi-context    Two independent search contexts on one page
      7. test-all-filters      Every filter type exercised (checkbox, daterange, slider, people, taxonomy, tagbox, toggle, dropdown, text)
      8. test-all-layouts      All 6 layouts enabled with layout switcher
      9. test-deep-link        Pre-configured URL with query + filters + sort + layout for deep-link testing
     10. test-hub-search       Hub/cross-site search with verticals
     11. test-no-filters       Results without filters web part (tests filter restoration timeout)
     12. test-manual-filters   Manual apply mode for filters

.PARAMETER SiteUrl
    Target SharePoint site URL.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP authentication.

.PARAMETER Publish
    Publish pages after creation. Defaults to true.

.PARAMETER PagesOnly
    Comma-separated list of page names to provision (e.g., "test-general,test-people").
    If omitted, all pages are provisioned.

.EXAMPLE
    # Provision all test pages
    .\Provision-TestPages.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/SPSearch" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f"

.EXAMPLE
    # Provision only specific pages
    .\Provision-TestPages.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/SPSearch" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f" `
      -PagesOnly "test-multi-context,test-all-filters"
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [bool]$Publish = $true,

    [Parameter(Mandatory = $false)]
    [string]$PagesOnly,

    # T4.D1 — bypass the destructive-op confirmation prompt for CI / scripted callers.
    # Without -Force, re-running over existing test pages prompts before recycling them.
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Web Part component names
# ============================================================================
$WP_SEARCH_BOX       = "SP Search Box"
$WP_SEARCH_RESULTS   = "SP Search Results"
$WP_SEARCH_FILTERS   = "SP Search Filters"
$WP_SEARCH_VERTICALS = "SP Search Verticals"
$WP_SEARCH_MANAGER   = "SP Search Manager"

# ============================================================================
# Prerequisites
# ============================================================================
$requiredModule = "PnP.PowerShell"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    throw "PnP.PowerShell module not found. Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
}
Import-Module $requiredModule -ErrorAction Stop

function Resolve-PnPClientId {
    param([string]$ExplicitClientId)
    if ($ExplicitClientId) { return $ExplicitClientId }
    foreach ($name in @('ENTRAID_APP_ID', 'ENTRAID_CLIENT_ID', 'AZURE_CLIENT_ID')) {
        $val = [Environment]::GetEnvironmentVariable($name)
        if (-not [string]::IsNullOrWhiteSpace($val)) { return $val.Trim() }
    }
    throw "Client ID required. Pass -ClientId or set ENTRAID_APP_ID / ENTRAID_CLIENT_ID / AZURE_CLIENT_ID."
}

function Add-SPSearchWebPart {
    param(
        [string]$Page,
        [string]$ComponentName,
        [int]$Section,
        [int]$Column,
        [int]$Order = 1,
        [hashtable]$Properties
    )
    Add-PnPPageWebPart -Page $Page `
        -Component $ComponentName `
        -Section $Section -Column $Column -Order $Order `
        -WebPartProperties $Properties `
        -ErrorAction Stop | Out-Null
}

function New-TestPage {
    param(
        [string]$PageName,
        [string]$PageTitle,
        [string]$Description
    )
    Write-Host "`n  Creating page: $PageName ($PageTitle)" -ForegroundColor Cyan
    Write-Host "    $Description" -ForegroundColor DarkGray

    # The top-level confirmation gate in the main flow already authorised the
    # recycle batch — see the "$existingTestPages.Count" ShouldProcess call below.
    try {
        Remove-PnPPage -Identity $PageName -Force -ErrorAction SilentlyContinue
    } catch { }

    Add-PnPPage -Name $PageName -Title $PageTitle -LayoutType Article -ErrorAction Stop | Out-Null
    return $PageName
}

# ============================================================================
# Connect
# ============================================================================
$resolvedClientId = Resolve-PnPClientId -ExplicitClientId $ClientId
Write-Host "`nConnecting to $SiteUrl ..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $resolvedClientId

$filterPages = @()
if ($PagesOnly) {
    $filterPages = $PagesOnly.Split(',') | ForEach-Object { $_.Trim() }
}

function Should-ProvisionPage {
    param([string]$Name)
    if ($filterPages.Count -eq 0) { return $true }
    return $filterPages -contains $Name
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host " SP Search Test Pages Provisioning" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

# T4.D1 — top-level safety gate. Enumerate which test pages already exist and
# require a single confirmation before recycling any of them. -Force bypasses
# the prompt; -WhatIf reports the recycle scope without executing.
$candidatePages = @(
    'test-general', 'test-documents', 'test-people', 'test-news', 'test-media',
    'test-multi-context', 'test-all-filters', 'test-all-layouts', 'test-hub-search',
    'test-no-filters', 'test-manual-filters', 'test-search-manager'
) | Where-Object { Should-ProvisionPage $_ }
$existingTestPages = @()
foreach ($name in $candidatePages) {
    $found = Get-PnPPage -Identity $name -ErrorAction SilentlyContinue
    if ($found) { $existingTestPages += $name }
}
if ($existingTestPages.Count -gt 0) {
    $target = "$($existingTestPages.Count) existing test page(s): " + ($existingTestPages -join ', ')
    if (-not ($Force -or $PSCmdlet.ShouldProcess($target, 'Recycle and recreate'))) {
        Write-Host "`n  Aborted — existing test pages left in place. Re-run with -Force to bypass the prompt, or -WhatIf to preview." -ForegroundColor Yellow
        return
    }
}

$pagesCreated = 0

# ============================================================================
# Page 1: General Search
# ============================================================================
if (Should-ProvisionPage "test-general") {
    $pg = New-TestPage -PageName "test-general" -PageTitle "Test: General Search" `
        -Description "Basic search — list layout, 3 filters, 4 verticals. Tests core search flow."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 2
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 3

    $ctx = "test-general"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search everything..."
        enableSuggestions = $true
        enableSearchManager = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_VERTICALS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        showCounts = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 3 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 10
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
        scenarioPreset = "general"
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 3 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 2: Document Search
# ============================================================================
if (Should-ProvisionPage "test-documents") {
    $pg = New-TestPage -PageName "test-documents" -PageTitle "Test: Document Search" `
        -Description "Document-scoped search — grid/compact/list layouts, file verticals, 4 sort options."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 2
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 3

    $ctx = "test-documents"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search documents..."
        enableSuggestions = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_VERTICALS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        showCounts = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 3 -Column 1 -Properties @{
        searchContextId = $ctx
        scenarioPreset = "documents"
        pageSize = 25
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
        enablePreviewPanel = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 3 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 3: People Search (Graph Provider)
# ============================================================================
if (Should-ProvisionPage "test-people") {
    $pg = New-TestPage -PageName "test-people" -PageTitle "Test: People Search" `
        -Description "Graph-backed people search — people layout, department/job/office filters."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-people"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search people..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        scenarioPreset = "people"
        pageSize = 20
        showPaging = $true
        showResultCount = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
    }
    $pagesCreated++
}

# ============================================================================
# Page 4: News Search
# ============================================================================
if (Should-ProvisionPage "test-news") {
    $pg = New-TestPage -PageName "test-news" -PageTitle "Test: News Search" `
        -Description "News articles — card layout, date-sorted, author/site/date filters."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-news"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search news..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        scenarioPreset = "news"
        pageSize = 12
        showPaging = $true
        showResultCount = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
    }
    $pagesCreated++
}

# ============================================================================
# Page 5: Media Gallery
# ============================================================================
if (Should-ProvisionPage "test-media") {
    $pg = New-TestPage -PageName "test-media" -PageTitle "Test: Media Gallery" `
        -Description "Image/video gallery — gallery layout, thumbnail optimization, image/video verticals."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-media"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search images and videos..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        scenarioPreset = "media"
        pageSize = 24
        showPaging = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
    }
    $pagesCreated++
}

# ============================================================================
# Page 6: Multi-Context (TWO independent searches on one page)
# ============================================================================
if (Should-ProvisionPage "test-multi-context") {
    $pg = New-TestPage -PageName "test-multi-context" -PageTitle "Test: Multi-Context Page" `
        -Description "TWO independent search experiences on one page. Tests searchContextId isolation + URL prefix namespacing."

    # Context A: Documents (top half)
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    # Context B: People (bottom half)
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 3
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 4

    $ctxA = "multi-docs"
    $ctxB = "multi-people"

    # Context A
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctxA
        placeholder = "Context A: Search documents..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctxA
        scenarioPreset = "documents"
        pageSize = 5
        showPaging = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctxA
        applyMode = "instant"
    }

    # Context B
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 3 -Column 1 -Properties @{
        searchContextId = $ctxB
        placeholder = "Context B: Search people..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 4 -Column 1 -Properties @{
        searchContextId = $ctxB
        scenarioPreset = "people"
        pageSize = 5
        showPaging = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 4 -Column 2 -Properties @{
        searchContextId = $ctxB
        applyMode = "instant"
    }
    $pagesCreated++
}

# ============================================================================
# Page 7: All Filter Types
# ============================================================================
if (Should-ProvisionPage "test-all-filters") {
    $pg = New-TestPage -PageName "test-all-filters" -PageTitle "Test: All Filter Types" `
        -Description "Every filter type exercised: checkbox, daterange, slider, people, taxonomy, tagbox, toggle, dropdown, text."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-all-filters"

    # Filters covering all 9 types using provisioned managed properties
    $allFilters = @(
        @{ managedProperty='FileType';            displayName='File Type';      filterType='checkbox';  urlAlias='ft'; showCount=$true; defaultExpanded=$true }
        @{ managedProperty='LastModifiedTime';    displayName='Modified Date';  filterType='daterange'; urlAlias='md'; defaultExpanded=$true }
        @{ managedProperty='AuthorOWSUSER';       displayName='Author';         filterType='people';    urlAlias='au'; defaultExpanded=$true }
        @{ managedProperty='SiteName';            displayName='Site';           filterType='tagbox';    urlAlias='si'; showCount=$true; defaultExpanded=$false }
        @{ managedProperty='ContentType';         displayName='Content Type';   filterType='dropdown';  urlAlias='ct'; showCount=$true; defaultExpanded=$false }
        @{ managedProperty='RefinableString07';   displayName='Is Active';      filterType='toggle';    urlAlias='ia'; trueLabel='Active'; falseLabel='Inactive'; defaultExpanded=$true }
        @{ managedProperty='RefinableDecimal00';  displayName='Budget Range';   filterType='slider';    urlAlias='bg'; defaultExpanded=$false }
        @{ managedProperty='RefinableString02';   displayName='Region';         filterType='checkbox';  urlAlias='rg'; showCount=$true; defaultExpanded=$false }
    ) | ConvertTo-Json -Depth 3 -Compress

    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search (testing all filter types)..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 10
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
        filtersCollection = $allFilters
    }
    $pagesCreated++
}

# ============================================================================
# Page 8: All Layouts
# ============================================================================
if (Should-ProvisionPage "test-all-layouts") {
    $pg = New-TestPage -PageName "test-all-layouts" -PageTitle "Test: All Layouts" `
        -Description "All 6 layouts enabled (list, compact, grid, card, people, gallery). Tests layout switching + chunk preloading."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-all-layouts"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search (try all 6 layouts)..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 12
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
        defaultLayout = "list"
        showListLayout = $true
        showCompactLayout = $true
        showGridLayout = $true
        showCardLayout = $true
        showPeopleLayout = $true
        showGalleryLayout = $true
        enablePreviewPanel = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 9: Hub / Cross-Site Search with Verticals
# ============================================================================
if (Should-ProvisionPage "test-hub-search") {
    $pg = New-TestPage -PageName "test-hub-search" -PageTitle "Test: Hub Search" `
        -Description "Cross-site search with All/Documents/News/Pages verticals. Tests per-vertical query templates."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 2
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 3

    $ctx = "test-hub"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search across all sites..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_VERTICALS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        showCounts = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 3 -Column 1 -Properties @{
        searchContextId = $ctx
        scenarioPreset = "hub-search"
        pageSize = 15
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
        searchScope = "All"
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 3 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 10: No Filters Web Part (tests URL filter restoration timeout)
# ============================================================================
if (Should-ProvisionPage "test-no-filters") {
    $pg = New-TestPage -PageName "test-no-filters" -PageTitle "Test: No Filters Web Part" `
        -Description "Results without Filters web part. Tests URL filter restoration 5s timeout + graceful degradation."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 2

    $ctx = "test-no-filters"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search (no filters on this page)..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 10
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 11: Manual Apply Filters
# ============================================================================
if (Should-ProvisionPage "test-manual-filters") {
    $pg = New-TestPage -PageName "test-manual-filters" -PageTitle "Test: Manual Apply Filters" `
        -Description "Filters in manual apply mode — select multiple, then click Apply. Tests pending state + Clear All."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2

    $ctx = "test-manual-filters"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search (manual filter apply)..."
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 10
        showPaging = $true
        showResultCount = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "manual"
        showClearAll = $true
    }
    $pagesCreated++
}

# ============================================================================
# Page 12: Search Manager (Standalone)
# ============================================================================
if (Should-ProvisionPage "test-search-manager") {
    $pg = New-TestPage -PageName "test-search-manager" -PageTitle "Test: Search Manager" `
        -Description "Search with standalone manager — saved searches, history, collections, health, insights."

    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 1
    Add-PnPPageSection -Page $pg -SectionTemplate TwoColumnLeft -Order 2
    Add-PnPPageSection -Page $pg -SectionTemplate OneColumn -Order 3

    $ctx = "test-manager"
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId = $ctx
        placeholder = "Search (with manager)..."
        enableSearchManager = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_RESULTS -Section 2 -Column 1 -Properties @{
        searchContextId = $ctx
        queryTemplate = "{searchTerms}"
        pageSize = 10
        showPaging = $true
        showResultCount = $true
        showSortDropdown = $true
        scenarioPreset = "general"
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_FILTERS -Section 2 -Column 2 -Properties @{
        searchContextId = $ctx
        applyMode = "instant"
        showClearAll = $true
    }
    Add-SPSearchWebPart -Page $pg -ComponentName $WP_SEARCH_MANAGER -Section 3 -Column 1 -Properties @{
        searchContextId = $ctx
        mode = "standalone"
        enableSavedSearches = $true
        enableSearchHistory = $true
        enableCollections = $true
    }
    $pagesCreated++
}

# ============================================================================
# Publish pages
# ============================================================================
if ($Publish -and $pagesCreated -gt 0) {
    Write-Host "`nPublishing $pagesCreated pages..." -ForegroundColor Yellow
    $allPages = @(
        "test-general", "test-documents", "test-people", "test-news",
        "test-media", "test-multi-context", "test-all-filters",
        "test-all-layouts", "test-hub-search", "test-no-filters",
        "test-manual-filters", "test-search-manager"
    )
    foreach ($pageName in $allPages) {
        if (Should-ProvisionPage $pageName) {
            try {
                Set-PnPPage -Identity $pageName -Publish -ErrorAction SilentlyContinue
                Write-Host "  Published: $pageName" -ForegroundColor Green
            } catch {
                Write-Host "  Publish skipped: $pageName ($($_.Exception.Message))" -ForegroundColor DarkYellow
            }
        }
    }
}

# ============================================================================
# Summary
# ============================================================================
Write-Host "`n========================================" -ForegroundColor Green
Write-Host " Test Pages Provisioned: $pagesCreated" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Test Checklist:" -ForegroundColor Cyan
Write-Host "  1. test-general         - Type query, verify results, filters, verticals"
Write-Host "  2. test-documents       - Switch layouts (grid/compact/list), sort, preview panel"
Write-Host "  3. test-people          - Verify Graph provider, people cards, presence"
Write-Host "  4. test-news            - Card layout, date sorting, thumbnail display"
Write-Host "  5. test-media           - Gallery layout, image zoom, video preview"
Write-Host "  6. test-multi-context   - Search in BOTH contexts independently"
Write-Host "  7. test-all-filters     - Test each filter type, Clear All, show more/less"
Write-Host "  8. test-all-layouts     - Switch between all 6 layouts, verify rendering"
Write-Host "  9. test-hub-search      - Vertical tabs, counts, per-vertical queries"
Write-Host " 10. test-no-filters      - Deep-link with filters param, verify 5s timeout"
Write-Host " 11. test-manual-filters  - Select filters, Apply button, pending state"
Write-Host " 12. test-search-manager  - Save search, view history, create collection"
Write-Host ""
Write-Host "Deep-Link Test URLs (append to site URL):" -ForegroundColor Cyan
Write-Host "  /SitePages/test-general.aspx?q=annual+report"
Write-Host "  /SitePages/test-documents.aspx?q=budget&ft=xlsx&s=LastModifiedTime:desc"
Write-Host "  /SitePages/test-all-layouts.aspx?q=project&l=grid"
Write-Host "  /SitePages/test-no-filters.aspx?q=test&ft=docx  (should timeout gracefully)"
Write-Host ""
