<#
.SYNOPSIS
    Resets the SP Search test site to a clean slate before reprovisioning.

.DESCRIPTION
    Removes custom libraries and lists created by the SP Search provisioning
    scripts, including:
    - Account Documents environment libraries
    - Generic Provision-TestData libraries and lists
    - SP Search hidden support lists

    Optionally removes provisioned search pages as well.
    This reset script does not remove term groups, term sets, terms,
    site columns, or content types.

.PARAMETER SiteUrl
    Target SharePoint site URL. Defaults to https://pixelboy.sharepoint.com/sites/SPSearch/

.PARAMETER ClientId
    Entra app client ID for PnP interactive auth. Defaults to 970bb320-0d49-4b4a-aa8f-c3f4b1e5928f

.PARAMETER RemovePages
    Also remove known provisioned search pages from Site Pages.

.EXAMPLE
    ./Reset-AccountDocumentsEnvironment.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/SPSearch" -ClientId "<app-id>"

.EXAMPLE
    ./Reset-AccountDocumentsEnvironment.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/SPSearch" -RemovePages
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl = 'https://pixelboy.sharepoint.com/sites/SPSearch/',

    [Parameter(Mandatory = $false)]
    [string]$ClientId = '970bb320-0d49-4b4a-aa8f-c3f4b1e5928f',

    [Parameter(Mandatory = $false)]
    [switch]$RemovePages
)

$ErrorActionPreference = 'Stop'

$ACCOUNT_DOCUMENT_LIBRARIES = @(
    'AccountDocs',
    'EntityDocs',
    'OperationalDocs',
    'ComplianceDocs',
    'ValuationDocs',
    'ArchiveDocs'
)

$GENERIC_TEST_LIBRARIES = @(
    'CorporatePolicies',
    'SalesMaterials',
    'MarketingContent',
    'HRResources',
    'FinanceReports',
    'EngineeringDocs',
    'LegalDocuments',
    'ProjectFiles',
    'MediaAssets',
    'KnowledgeBase'
)

$GENERIC_TEST_LISTS = @(
    'Projects',
    'Contacts',
    'Tasks',
    'Events',
    'Inventory',
    'Announcements',
    'Issues',
    'FAQ',
    'Policies',
    'Glossary'
)

$SEARCH_SUPPORT_LISTS = @(
    'SearchSavedQueries',
    'SearchHistory',
    'SearchCollections'
)

$KNOWN_SEARCH_PAGES = @(
    'AccountDocumentsSearch',
    'Search',
    'Search1'
)

function Resolve-PnPClientId {
    param(
        [string]$ExplicitClientId
    )

    if ($ExplicitClientId) {
        return $ExplicitClientId
    }

    $candidateNames = @(
        'ENTRAID_APP_ID',
        'ENTRAID_CLIENT_ID',
        'AZURE_CLIENT_ID'
    )

    foreach ($candidateName in $candidateNames) {
        $candidateValue = [Environment]::GetEnvironmentVariable($candidateName)
        if (-not [string]::IsNullOrWhiteSpace($candidateValue)) {
            return $candidateValue.Trim()
        }
    }

    throw 'PnP interactive auth now requires an Entra app client ID. Re-run with -ClientId <app-id>, or set one of these environment variables before running: ENTRAID_APP_ID, ENTRAID_CLIENT_ID, AZURE_CLIENT_ID.'
}

function Normalize-SiteUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Trim().TrimEnd('/').ToLowerInvariant()
}

function Use-PnPConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetSiteUrl,

        [string]$ExplicitClientId
    )

    $connectionState = @{
        DisconnectOnExit = $false
    }

    $existingConnection = $null
    try {
        $existingConnection = Get-PnPConnection -ErrorAction Stop
    } catch {
        $existingConnection = $null
    }

    $normalizedTargetSiteUrl = Normalize-SiteUrl -Value $TargetSiteUrl
    $currentConnectionUrl = ''
    if ($existingConnection -and $existingConnection.Url) {
        $currentConnectionUrl = Normalize-SiteUrl -Value $existingConnection.Url
    }

    if ($existingConnection -and $currentConnectionUrl -eq $normalizedTargetSiteUrl) {
        Write-Host '  Reusing existing PnP connection' -ForegroundColor Green
        return $connectionState
    }

    $resolvedClientId = Resolve-PnPClientId -ExplicitClientId $ExplicitClientId
    Connect-PnPOnline -Url $TargetSiteUrl -ClientId $resolvedClientId -Interactive
    $connectionState.DisconnectOnExit = $true
    Write-Host "  Connected to $TargetSiteUrl" -ForegroundColor Green
    return $connectionState
}

function Get-PageServerRelativeUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseSiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$TargetPageName
    )

    $siteUri = [Uri]$BaseSiteUrl
    $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = ''
    }

    return "$sitePath/SitePages/$TargetPageName.aspx"
}

function Remove-ListIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if (-not $list) {
        return $false
    }

    Remove-PnPList -Identity $ListName -Force -ErrorAction Stop
    return $true
}

function Remove-PageIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseSiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$PageName
    )

    $pageServerRelativeUrl = Get-PageServerRelativeUrl -BaseSiteUrl $BaseSiteUrl -TargetPageName $PageName

    try {
        Get-PnPFile -Url $pageServerRelativeUrl -AsListItem -ErrorAction Stop | Out-Null
    } catch {
        return $false
    }

    try {
        Undo-PnPFileCheckedOut -Url $pageServerRelativeUrl -ErrorAction Stop
    } catch { }

    try {
        Remove-PnPFile -ServerRelativeUrl $pageServerRelativeUrl -Force -Recycle -ErrorAction Stop
    } catch {
        Remove-PnPPage -Identity $PageName -Force -ErrorAction SilentlyContinue
    }

    return $true
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ' SP Search Test Site Reset' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Site:        $SiteUrl" -ForegroundColor White
Write-Host "  RemovePages: $($RemovePages.IsPresent)" -ForegroundColor White
Write-Host ''

$connectionState = $null
$removed = New-Object System.Collections.Generic.List[string]

try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    $connectionState = Use-PnPConnection -TargetSiteUrl $SiteUrl -ExplicitClientId $ClientId

    $allListsToRemove = @(
        $ACCOUNT_DOCUMENT_LIBRARIES +
        $GENERIC_TEST_LIBRARIES +
        $GENERIC_TEST_LISTS +
        $SEARCH_SUPPORT_LISTS
    ) | Select-Object -Unique

    Write-Host '  Removing provisioned libraries and lists...' -ForegroundColor Gray
    foreach ($listName in $allListsToRemove) {
        try {
            if (Remove-ListIfExists -ListName $listName) {
                Write-Host "    Removed '$listName'" -ForegroundColor Yellow
                $removed.Add($listName)
            }
        } catch {
            Write-Warning "    Failed to remove '$listName': $($_.Exception.Message)"
        }
    }

    if ($RemovePages) {
        Write-Host '  Removing provisioned search pages...' -ForegroundColor Gray
        foreach ($pageName in $KNOWN_SEARCH_PAGES) {
            try {
                if (Remove-PageIfExists -BaseSiteUrl $SiteUrl -PageName $pageName) {
                    Write-Host "    Removed '$pageName.aspx'" -ForegroundColor Yellow
                    $removed.Add("$pageName.aspx")
                }
            } catch {
                Write-Warning "    Failed to remove '$pageName.aspx': $($_.Exception.Message)"
            }
        }
    }

    Write-Host ''
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ' Reset complete' -ForegroundColor Green
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ''
    Write-Host "  Removed objects: $($removed.Count)" -ForegroundColor White
    if ($removed.Count -gt 0) {
        $removed | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Gray
        }
    } else {
        Write-Host '  Nothing to remove. Site was already clean for these known artifacts.' -ForegroundColor Gray
    }
    Write-Host ''
} catch {
    Write-Host ''
    Write-Host '======================================================================' -ForegroundColor Red
    Write-Host ' Reset failed' -ForegroundColor Red
    Write-Host '======================================================================' -ForegroundColor Red
    Write-Host ''
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    Write-Host ''
    Write-Host 'Stack:' -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ''
    exit 1
} finally {
    if ($connectionState -and $connectionState.DisconnectOnExit) {
        try {
            $null = Get-PnPConnection -ErrorAction Stop
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Host 'Disconnected from SharePoint' -ForegroundColor Gray
        } catch { }
    }
}
