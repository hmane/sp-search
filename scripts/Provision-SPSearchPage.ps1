<#
.SYNOPSIS
    Provisions a SharePoint client-side search page with all SP Search web parts.

.DESCRIPTION
    Creates a modern SharePoint page with the optimal search layout:
    - Section 1 (full width): Search Box
    - Section 2 (full width): Verticals (tab navigation)
    - Section 3 (two-column 66/33): Results (left) | Filters (right)

    All web parts share the same searchContextId for state synchronization.
    Optionally sets the page as the site home page.

    Idempotent — re-running overwrites the existing page.

.PARAMETER SiteUrl
    Target SharePoint site URL.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP authentication.

.PARAMETER PageName
    Page file name (without .aspx). Defaults to "Search".

.PARAMETER PageTitle
    Page display title. Defaults to "Search".

.PARAMETER SearchContextId
    Shared context ID for all web parts on the page. Defaults to "default".

.PARAMETER SetAsHomePage
    Set the search page as the site's home page.

.PARAMETER IncludeSearchManager
    Add a Search Manager web part below the results area (standalone mode).

.PARAMETER Publish
    Publish the page after creation. Defaults to true.

.EXAMPLE
    # Basic search page
    .\Provision-SPSearchPage.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/search" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f"

.EXAMPLE
    # Search center with custom name, set as home page
    .\Provision-SPSearchPage.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/search" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f" `
      -PageName "SearchCenter" `
      -PageTitle "Search Center" `
      -SetAsHomePage `
      -IncludeSearchManager

.EXAMPLE
    # Multi-context: create a second search page with isolated state
    .\Provision-SPSearchPage.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/hr" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f" `
      -PageName "PeopleSearch" `
      -PageTitle "People Search" `
      -SearchContextId "people"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Target SharePoint site URL")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false, HelpMessage = "Azure AD Client ID for PnP authentication")]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$PageName = "Search",

    [Parameter(Mandatory = $false)]
    [string]$PageTitle = "Search",

    [Parameter(Mandatory = $false)]
    [string]$SearchContextId = "default",

    [Parameter(Mandatory = $false)]
    [switch]$SetAsHomePage,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSearchManager,

    [Parameter(Mandatory = $false)]
    [bool]$Publish = $true
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Web Part Component IDs (from manifest files)
# ============================================================================
$WP_SEARCH_BOX       = "13a82dbe-2c57-4e20-bfe8-ec4de5776191"
$WP_SEARCH_RESULTS   = "1836671c-a710-45b4-9a83-55c65344a3d5"
$WP_SEARCH_FILTERS   = "2eb68250-879f-45a8-af9b-9fc3e97b2050"
$WP_SEARCH_VERTICALS = "d0481c49-49f9-4219-90fe-be8338051f58"
$WP_SEARCH_MANAGER   = "46308c1c-af6b-43c5-98b7-2d39082498cb"

# ============================================================================
# Default vertical configuration
# ============================================================================
$defaultVerticals = @(
    @{
        key           = "all"
        label         = "All"
        iconName      = "Search"
        sortOrder     = 1
    },
    @{
        key           = "documents"
        label         = "Documents"
        iconName      = "Page"
        queryTemplate = "{searchTerms} contentclass:STS_ListItem_DocumentLibrary"
        sortOrder     = 2
    },
    @{
        key           = "pages"
        label         = "Pages"
        iconName      = "FileHTML"
        queryTemplate = "{searchTerms} (contentclass:STS_ListItem_WebPageLibrary OR contentclass:STS_Site)"
        sortOrder     = 3
    },
    @{
        key           = "people"
        label         = "People"
        iconName      = "People"
        queryTemplate = "{searchTerms}"
        resultSourceId = "b09a7990-05ea-4af9-81ef-edfab16c4e31"
        sortOrder     = 4
    },
    @{
        key           = "sites"
        label         = "Sites"
        iconName      = "Globe"
        queryTemplate = "{searchTerms} contentclass:STS_Site"
        sortOrder     = 5
    }
) | ConvertTo-Json -Depth 4 -Compress

# ============================================================================
# Validate prerequisites
# ============================================================================
$requiredModule = "PnP.PowerShell"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    throw "PnP.PowerShell module not found. Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
}

Import-Module $requiredModule -ErrorAction Stop

# ============================================================================
# Helper: Add SPFx web part to a page
# ============================================================================
function Add-SPSearchWebPart {
    param(
        [string]$Page,
        [string]$ComponentId,
        [int]$Section,
        [int]$Column,
        [int]$Order = 1,
        [hashtable]$Properties
    )

    $propsJson = $Properties | ConvertTo-Json -Depth 10 -Compress

    Add-PnPPageWebPart -Page $Page `
        -Component $ComponentId `
        -Section $Section `
        -Column $Column `
        -Order $Order `
        -WebPartProperties $Properties `
        -ErrorAction Stop | Out-Null
}

# ============================================================================
# Main
# ============================================================================
Write-Host ""
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host " SP Search — Page Provisioning" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Site:      $SiteUrl" -ForegroundColor White
Write-Host "  Page:      $PageName.aspx" -ForegroundColor White
Write-Host "  Title:     $PageTitle" -ForegroundColor White
Write-Host "  Context:   $SearchContextId" -ForegroundColor White
if ($SetAsHomePage) {
    Write-Host "  Home page: Yes" -ForegroundColor White
}
if ($IncludeSearchManager) {
    Write-Host "  Manager:   Standalone (below results)" -ForegroundColor White
}
Write-Host ""

$totalSteps = 5
if ($SetAsHomePage) { $totalSteps++ }
if ($IncludeSearchManager) { $totalSteps++ }
$step = 0

try {
    # ─── Step 1: Connect ──────────────────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Connecting to SharePoint..." -ForegroundColor Cyan
    if ($ClientId) {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
    } else {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }
    Write-Host "  Connected successfully" -ForegroundColor Green
    Write-Host ""

    # ─── Step 2: Create page ──────────────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Creating page '$PageName.aspx'..." -ForegroundColor Cyan

    # Check if page exists
    $existingPage = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue
    if ($existingPage) {
        Write-Host "  [EXISTS] Page '$PageName.aspx' exists — removing for clean recreation..." -ForegroundColor Yellow
        Remove-PnPPage -Identity $PageName -Force -ErrorAction Stop
    }

    # Create the page with Article layout (standard content page)
    Add-PnPPage -Name $PageName -Title $PageTitle -LayoutType Article -HeaderLayoutType NoImage -CommentsEnabled:$false -ErrorAction Stop | Out-Null
    Write-Host "  Page created" -ForegroundColor Green
    Write-Host ""

    # ─── Step 3: Add page sections ────────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Configuring page layout..." -ForegroundColor Cyan

    # Section 1: Full width — Search Box
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 1 -ErrorAction Stop
    Write-Host "  Section 1: OneColumn (Search Box)" -ForegroundColor Green

    # Section 2: Full width — Verticals
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 2 -ErrorAction Stop
    Write-Host "  Section 2: OneColumn (Verticals)" -ForegroundColor Green

    # Section 3: Two-column (66% left / 33% right) — Results | Filters
    Add-PnPPageSection -Page $PageName -SectionTemplate TwoColumnLeft -Order 3 -ErrorAction Stop
    Write-Host "  Section 3: TwoColumnLeft (Results | Filters)" -ForegroundColor Green

    if ($IncludeSearchManager) {
        # Section 4: Full width — Search Manager
        Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 4 -ErrorAction Stop
        Write-Host "  Section 4: OneColumn (Search Manager)" -ForegroundColor Green
    }
    Write-Host ""

    # ─── Step 4: Add web parts ────────────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Adding SP Search web parts..." -ForegroundColor Cyan

    # Search Box (Section 1)
    Write-Host "  Adding Search Box..." -ForegroundColor Yellow
    Add-SPSearchWebPart -Page $PageName -ComponentId $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId     = $SearchContextId
        placeholder         = "Search SharePoint..."
        debounceMs          = 300
        searchBehavior      = "both"
        enableScopeSelector = $true
        enableSuggestions    = $true
        enableSearchManager = (-not $IncludeSearchManager)
        enableQueryBuilder  = $false
    }
    Write-Host "  [OK] Search Box" -ForegroundColor Green

    # Verticals (Section 2)
    Write-Host "  Adding Verticals..." -ForegroundColor Yellow
    Add-SPSearchWebPart -Page $PageName -ComponentId $WP_SEARCH_VERTICALS -Section 2 -Column 1 -Properties @{
        searchContextId    = $SearchContextId
        verticals          = $defaultVerticals
        showCounts         = $true
        hideEmptyVerticals = $false
        tabStyle           = "underline"
    }
    Write-Host "  [OK] Verticals (All, Documents, Pages, People, Sites)" -ForegroundColor Green

    # Search Results (Section 3, Column 1 — left, wider)
    Write-Host "  Adding Search Results..." -ForegroundColor Yellow
    Add-SPSearchWebPart -Page $PageName -ComponentId $WP_SEARCH_RESULTS -Section 3 -Column 1 -Properties @{
        searchContextId  = $SearchContextId
        pageSize         = 25
        defaultLayout    = "list"
        showResultCount  = $true
        showSortDropdown = $true
        enableSelection  = $true
    }
    Write-Host "  [OK] Search Results (List layout, 25 per page)" -ForegroundColor Green

    # Search Filters (Section 3, Column 2 — right, narrower)
    Write-Host "  Adding Search Filters..." -ForegroundColor Yellow
    Add-SPSearchWebPart -Page $PageName -ComponentId $WP_SEARCH_FILTERS -Section 3 -Column 2 -Properties @{
        searchContextId        = $SearchContextId
        applyMode              = "instant"
        operatorBetweenFilters = "AND"
        showClearAll           = $true
    }
    Write-Host "  [OK] Search Filters (instant apply, AND operator)" -ForegroundColor Green

    Write-Host ""

    # ─── Step 4b: Search Manager (optional) ───────────────────
    if ($IncludeSearchManager) {
        $step++
        Write-Host "[$step/$totalSteps] Adding Search Manager..." -ForegroundColor Cyan
        Add-SPSearchWebPart -Page $PageName -ComponentId $WP_SEARCH_MANAGER -Section 4 -Column 1 -Properties @{
            searchContextId = $SearchContextId
            mode            = "standalone"
        }
        Write-Host "  [OK] Search Manager (standalone mode)" -ForegroundColor Green
        Write-Host ""
    }

    # ─── Step 5: Publish ──────────────────────────────────────
    $step++
    if ($Publish) {
        Write-Host "[$step/$totalSteps] Publishing page..." -ForegroundColor Cyan
        Set-PnPPage -Identity $PageName -Publish -ErrorAction Stop
        Write-Host "  Page published" -ForegroundColor Green
    } else {
        Write-Host "[$step/$totalSteps] Skipping publish (page saved as draft)" -ForegroundColor Yellow
    }
    Write-Host ""

    # ─── Step 6: Set as home page (optional) ──────────────────
    if ($SetAsHomePage) {
        $step++
        Write-Host "[$step/$totalSteps] Setting as site home page..." -ForegroundColor Cyan
        Set-PnPHomePage -RootFolderRelativeUrl "SitePages/$PageName.aspx" -ErrorAction Stop
        Write-Host "  Home page updated to $PageName.aspx" -ForegroundColor Green
        Write-Host ""
    }

    # ─── Summary ──────────────────────────────────────────────
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host " Page provisioning complete!" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Page URL:  $SiteUrl/SitePages/$PageName.aspx" -ForegroundColor White
    Write-Host "  Context:   $SearchContextId" -ForegroundColor White
    Write-Host ""
    Write-Host "  Layout:" -ForegroundColor Yellow
    Write-Host "  ┌──────────────────────────────────────────────────────┐"
    Write-Host "  │ Search Box (full width)                              │"
    Write-Host "  ├──────────────────────────────────────────────────────┤"
    Write-Host "  │ Verticals: All │ Documents │ Pages │ People │ Sites │"
    Write-Host "  ├────────────────────────────┬─────────────────────────┤"
    Write-Host "  │ Results (66%)              │ Filters (33%)          │"
    Write-Host "  │  - List layout, 25/page    │  - Instant apply       │"
    Write-Host "  │  - Sort + selection        │  - AND operator        │"
    if ($IncludeSearchManager) {
        Write-Host "  ├────────────────────────────┴─────────────────────────┤"
        Write-Host "  │ Search Manager (saved searches, collections, history)│"
    }
    Write-Host "  └──────────────────────────────────────────────────────┘"
    Write-Host ""
    Write-Host "  Customize web parts by editing the page in SharePoint." -ForegroundColor Gray
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Red
    Write-Host " Page provisioning failed at step $step!" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red

    if ($_.Exception.InnerException) {
        Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "Stack:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ""
    exit 1

} finally {
    try {
        $null = Get-PnPConnection -ErrorAction Stop
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
    } catch {
        # Already disconnected
    }
}
