<#
.SYNOPSIS
    Exports SP Search web part configuration from a modern SharePoint page.

.DESCRIPTION
    Reads the page canvas and exports the raw SPFx properties for every SP Search
    web part on the page. The export intentionally preserves the complete
    properties object instead of maintaining a whitelist, so new web part
    properties are included automatically.

    Use Import-SPSearchPageConfig.ps1 to apply the exported JSON to a matching
    page in another environment.

.PARAMETER SiteUrl
    Source SharePoint site URL.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP interactive authentication.
    If omitted, the script reads ENTRAID_APP_ID, ENTRAID_CLIENT_ID, or
    AZURE_CLIENT_ID.

.PARAMETER PageName
    Source page file name, with or without .aspx. Defaults to Search.aspx.

.PARAMETER OutputPath
    JSON file to write. Defaults to .\sp-search-page-config.<page>.json.

.PARAMETER TokenizeSiteUrl
    Replaces the source site URL inside exported string values with {siteUrl}.
    This is useful when the JSON is committed and imported into another site.

.PARAMETER Force
    Overwrite OutputPath if it already exists.

.EXAMPLE
    .\scripts\Export-SPSearchPageConfig.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/search-dev" `
        -ClientId "<app-id>" `
        -PageName "Search" `
        -OutputPath ".\config\search-page.dev.json" `
        -TokenizeSiteUrl `
        -Force
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingWriteHost", "", Justification = "Provisioning scripts use colored host output for operator progress.")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseApprovedVerbs", "", Justification = "Private script helpers are not exported cmdlets.")]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification = "Private script helpers use domain language.")]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$PageName = "Search",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$TokenizeSiteUrl,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$requiredModule = "PnP.PowerShell"
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    throw "PnP.PowerShell module not found. Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
}

Import-Module $requiredModule -ErrorAction Stop

$script:SpSearchComponentIds = @{
    "13a82dbe-2c57-4e20-bfe8-ec4de5776191" = "SP Search Box"
    "1836671c-a710-45b4-9a83-55c65344a3d5" = "SP Search Results"
    "2eb68250-879f-45a8-af9b-9fc3e97b2050" = "SP Search Filters"
    "d0481c49-49f9-4219-90fe-be8338051f58" = "SP Search Verticals"
    "46308c1c-af6b-43c5-98b7-2d39082498cb" = "SP Search Manager"
    "17007020-148e-49b8-a628-972fa08139c6" = "SP Search Admin Manager"
}

function Resolve-PnPClientId {
    param(
        [string]$ExplicitClientId
    )

    if ($ExplicitClientId) {
        return $ExplicitClientId
    }

    $candidateNames = @(
        "ENTRAID_APP_ID",
        "ENTRAID_CLIENT_ID",
        "AZURE_CLIENT_ID"
    )

    foreach ($candidateName in $candidateNames) {
        $candidateValue = [Environment]::GetEnvironmentVariable($candidateName)
        if (-not [string]::IsNullOrWhiteSpace($candidateValue)) {
            return $candidateValue.Trim()
        }
    }

    throw "PnP interactive auth requires an Entra app client ID. Re-run with -ClientId <app-id>, or set ENTRAID_APP_ID, ENTRAID_CLIENT_ID, or AZURE_CLIENT_ID."
}

function Normalize-SiteUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Trim().TrimEnd("/")
}

function Normalize-PageLeafName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $leafName = [IO.Path]::GetFileName($Value.Trim())
    if (-not $leafName.EndsWith(".aspx", [StringComparison]::OrdinalIgnoreCase)) {
        $leafName = "$leafName.aspx"
    }

    return $leafName
}

function ConvertTo-JsonCompatible {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int] -or
        $Value -is [int64] -or
        $Value -is [decimal] -or
        $Value -is [double] -or
        $Value -is [single] -or
        $Value -is [datetime] -or
        $Value -is [guid]) {
        return $Value
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $map = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $map[$key] = ConvertTo-JsonCompatible -Value $Value[$key]
        }
        return $map
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @()
        foreach ($item in $Value) {
            $items += ,(ConvertTo-JsonCompatible -Value $item)
        }
        return $items
    }

    $objectMap = [ordered]@{}
    foreach ($property in $Value.PSObject.Properties) {
        $objectMap[$property.Name] = ConvertTo-JsonCompatible -Value $property.Value
    }

    return $objectMap
}

function Get-MapValue {
    param(
        [AllowNull()]
        [object]$Map,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    if ($null -eq $Map) {
        return $null
    }

    if ($Map -is [System.Collections.IDictionary] -and $Map.Contains($Key)) {
        return $Map[$Key]
    }

    return $null
}

function Normalize-SpSearchComponentName {
    param(
        [AllowNull()]
        [string]$Name,

        [AllowNull()]
        [string]$ComponentId
    )

    if (-not [string]::IsNullOrWhiteSpace($ComponentId)) {
        $componentIdKey = $ComponentId.ToLowerInvariant()
        if ($script:SpSearchComponentIds.ContainsKey($componentIdKey)) {
            return $script:SpSearchComponentIds[$componentIdKey]
        }
    }

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $normalized = ($Name -replace "\s", "").ToLowerInvariant()
    switch ($normalized) {
        "spsearchbox" { return "SP Search Box" }
        "spsearchresults" { return "SP Search Results" }
        "spsearchfilters" { return "SP Search Filters" }
        "spsearchverticals" { return "SP Search Verticals" }
        "spsearchmanager" { return "SP Search Manager" }
        "spsearchadminmanager" { return "SP Search Admin Manager" }
        default { return $null }
    }
}

function Replace-InJsonCompatibleObject {
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Find,

        [Parameter(Mandatory = $true)]
        [string]$ReplaceWith
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        return $Value.Replace($Find, $ReplaceWith)
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $map = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $map[$key] = Replace-InJsonCompatibleObject -Value $Value[$key] -Find $Find -ReplaceWith $ReplaceWith
        }
        return $map
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @()
        foreach ($item in $Value) {
            $items += ,(Replace-InJsonCompatibleObject -Value $item -Find $Find -ReplaceWith $ReplaceWith)
        }
        return $items
    }

    return $Value
}

function Get-PageListItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeafName
    )

    $escapedLeafName = $LeafName.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
    $query = @"
<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='FileLeafRef' />
        <Value Type='File'>$escapedLeafName</Value>
      </Eq>
    </Where>
  </Query>
  <RowLimit>1</RowLimit>
</View>
"@

    $candidateLists = @("Site Pages", "SitePages")
    foreach ($listName in $candidateLists) {
        try {
            $items = Get-PnPListItem -List $listName -Query $query -Fields "FileLeafRef", "FileRef", "CanvasContent1" -ErrorAction Stop
            $item = $items | Select-Object -First 1
            if ($item) {
                return [ordered]@{
                    ListName = $listName
                    Item = $item
                }
            }
        } catch {
            Write-Verbose "Could not read '$LeafName' from '$listName': $($_.Exception.Message)"
        }
    }

    throw "Page '$LeafName' was not found in Site Pages."
}

function Get-SpSearchWebPartsFromCanvas {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Canvas
    )

    $webParts = @()
    $countsByComponent = @{}

    foreach ($control in $Canvas) {
        $data = Get-MapValue -Map $control -Key "data"
        $webPartData = Get-MapValue -Map $data -Key "webPartData"
        if ($null -eq $webPartData) {
            continue
        }

        $componentId = Get-MapValue -Map $webPartData -Key "id"
        if ([string]::IsNullOrWhiteSpace($componentId)) {
            $componentId = Get-MapValue -Map $data -Key "webPartId"
        }
        $title = Get-MapValue -Map $webPartData -Key "title"
        $componentName = Normalize-SpSearchComponentName -Name $title -ComponentId $componentId
        if (-not $componentName) {
            continue
        }

        if (-not $countsByComponent.ContainsKey($componentName)) {
            $countsByComponent[$componentName] = 0
        }
        $countsByComponent[$componentName]++

        $properties = Get-MapValue -Map $webPartData -Key "properties"
        if ($null -eq $properties) {
            $properties = [ordered]@{}
        }

        $instanceId = Get-MapValue -Map $control -Key "id"
        if ([string]::IsNullOrWhiteSpace($instanceId)) {
            $instanceId = Get-MapValue -Map $webPartData -Key "instanceId"
        }

        $webParts += ,[ordered]@{
            key = "{0}:{1}" -f ($componentName -replace "\s", "").ToLowerInvariant(), $countsByComponent[$componentName]
            component = $componentName
            title = $title
            componentId = $componentId
            instanceId = $instanceId
            occurrence = $countsByComponent[$componentName]
            searchContextId = Get-MapValue -Map $properties -Key "searchContextId"
            position = Get-MapValue -Map $control -Key "position"
            properties = $properties
            serverProcessedContent = Get-MapValue -Map $webPartData -Key "serverProcessedContent"
            dynamicDataPaths = Get-MapValue -Map $webPartData -Key "dynamicDataPaths"
            dynamicDataValues = Get-MapValue -Map $webPartData -Key "dynamicDataValues"
        }
    }

    return $webParts
}

$normalizedSiteUrl = Normalize-SiteUrl -Value $SiteUrl
$leafName = Normalize-PageLeafName -Value $PageName
$pageNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($leafName)

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Get-Location) "sp-search-page-config.$pageNameWithoutExtension.json"
}

$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
    throw "Output file already exists: $resolvedOutputPath. Re-run with -Force to overwrite it."
}

Write-Host ""
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host " SP Search - Export Page Configuration" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Site:   $normalizedSiteUrl" -ForegroundColor White
Write-Host "  Page:   $leafName" -ForegroundColor White
Write-Host "  Output: $resolvedOutputPath" -ForegroundColor White
if ($TokenizeSiteUrl) {
    Write-Host "  Token:  replacing site URL with {siteUrl}" -ForegroundColor White
}
Write-Host ""

$existingConnection = $null
try {
    $existingConnection = Get-PnPConnection -ErrorAction Stop
} catch {
    $existingConnection = $null
}

$currentConnectionUrl = ""
if ($existingConnection -and $existingConnection.Url) {
    $currentConnectionUrl = Normalize-SiteUrl -Value $existingConnection.Url
}

if ($existingConnection -and $currentConnectionUrl -eq $normalizedSiteUrl) {
    Write-Host "Reusing existing PnP connection" -ForegroundColor Green
} else {
    $resolvedClientId = Resolve-PnPClientId -ExplicitClientId $ClientId
    Connect-PnPOnline -Url $normalizedSiteUrl -ClientId $resolvedClientId -Interactive
    Write-Host "Connected to SharePoint" -ForegroundColor Green
}

$pageInfo = Get-PageListItem -LeafName $leafName
$item = $pageInfo.Item
$canvasRaw = $item["CanvasContent1"]
if ([string]::IsNullOrWhiteSpace($canvasRaw)) {
    throw "Page '$leafName' does not contain CanvasContent1 data."
}

$canvas = @(ConvertTo-JsonCompatible -Value (ConvertFrom-Json -InputObject $canvasRaw))
$webParts = @(Get-SpSearchWebPartsFromCanvas -Canvas $canvas)
if ($webParts.Count -eq 0) {
    throw "No SP Search web parts were found on '$leafName'."
}

if ($TokenizeSiteUrl) {
    for ($index = 0; $index -lt $webParts.Count; $index++) {
        $webParts[$index].properties = Replace-InJsonCompatibleObject -Value $webParts[$index].properties -Find $normalizedSiteUrl -ReplaceWith "{siteUrl}"
        $webParts[$index].serverProcessedContent = Replace-InJsonCompatibleObject -Value $webParts[$index].serverProcessedContent -Find $normalizedSiteUrl -ReplaceWith "{siteUrl}"
        $webParts[$index].dynamicDataPaths = Replace-InJsonCompatibleObject -Value $webParts[$index].dynamicDataPaths -Find $normalizedSiteUrl -ReplaceWith "{siteUrl}"
        $webParts[$index].dynamicDataValues = Replace-InJsonCompatibleObject -Value $webParts[$index].dynamicDataValues -Find $normalizedSiteUrl -ReplaceWith "{siteUrl}"
    }
}

$export = [ordered]@{
    '$schema' = "./sp-search-page-config.schema.json"
    schemaVersion = 1
    solution = "sp-search"
    exportedAt = (Get-Date).ToUniversalTime().ToString("o")
    source = [ordered]@{
        siteUrl = $(if ($TokenizeSiteUrl) { "{siteUrl}" } else { $normalizedSiteUrl })
        pageName = $leafName
        pageUrl = "$(if ($TokenizeSiteUrl) { "{siteUrl}" } else { $normalizedSiteUrl })/SitePages/$leafName"
        listName = $pageInfo.ListName
        itemId = $item.Id
    }
    importNotes = [ordered]@{
        matching = "Import matches by instanceId first, then component + searchContextId + occurrence, then component + occurrence."
        scope = "Only SP Search web part property bags are exported. Page sections, non-SP Search controls, list data, and per-user local state are not migrated."
        tokenReplacement = "Use {siteUrl} for target-site-specific URLs. Import always provides {siteUrl}; TokenFile can provide additional tokens."
    }
    webParts = $webParts
}

$outputDirectory = Split-Path -Parent $resolvedOutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$json = $export | ConvertTo-Json -Depth 100
Set-Content -LiteralPath $resolvedOutputPath -Value $json -Encoding UTF8

Write-Host ""
Write-Host "Exported $($webParts.Count) SP Search web part configuration(s):" -ForegroundColor Green
foreach ($webPart in $webParts) {
    $context = $webPart.searchContextId
    if ([string]::IsNullOrWhiteSpace($context)) {
        $context = "(no context)"
    }
    Write-Host ("  - {0} occurrence {1}, context {2}" -f $webPart.component, $webPart.occurrence, $context) -ForegroundColor White
}
Write-Host ""
Write-Host "Wrote $resolvedOutputPath" -ForegroundColor Green
