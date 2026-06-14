<#
.SYNOPSIS
    Imports SP Search web part configuration into a modern SharePoint page.

.DESCRIPTION
    Applies a JSON export produced by Export-SPSearchPageConfig.ps1 to an
    existing page. The script updates the raw SPFx properties for matching SP
    Search web parts on the target page and preserves page layout, sections,
    non-SP Search controls, and existing target web part instance IDs.

    Matching order:
    1. exported instanceId -> target canvas control id
    2. component + searchContextId + occurrence
    3. component + occurrence
    4. component + searchContextId, only when exactly one target candidate exists

    The target page should already contain the SP Search web parts. Use
    Provision-SPSearchPage.ps1 or Search-ScenarioPresets.ps1 to create the page
    first, then import the saved configuration.

.PARAMETER SiteUrl
    Target SharePoint site URL.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP interactive authentication.
    If omitted, the script reads ENTRAID_APP_ID, ENTRAID_CLIENT_ID, or
    AZURE_CLIENT_ID.

.PARAMETER PageName
    Target page file name, with or without .aspx. Defaults to Search.aspx.

.PARAMETER ConfigPath
    JSON export to import.

.PARAMETER TokenFile
    Optional JSON object with token values. Tokens can be flat:
    { "managedPropertyPrefix": "RefinableString" }
    or wrapped:
    { "tokens": { "managedPropertyPrefix": "RefinableString" } }

    Import always adds {siteUrl} for the target site URL.

.PARAMETER Publish
    Publishes the page after updating CanvasContent1. Defaults to true.

.PARAMETER Force
    Applies changes without the confirmation prompt. -WhatIf still previews
    without writing.

.EXAMPLE
    .\scripts\Import-SPSearchPageConfig.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/search-prod" `
        -ClientId "<app-id>" `
        -PageName "Search" `
        -ConfigPath ".\config\search-page.dev.json" `
        -Force

.EXAMPLE
    .\scripts\Import-SPSearchPageConfig.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/search-prod" `
        -ConfigPath ".\config\search-page.dev.json" `
        -TokenFile ".\config\prod.tokens.json" `
        -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
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

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$TokenFile,

    [Parameter(Mandatory = $false)]
    [bool]$Publish = $true,

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

function Apply-Tokens {
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Tokens
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        $result = $Value
        foreach ($tokenName in $Tokens.Keys) {
            $tokenValue = [string]$Tokens[$tokenName]
            $result = $result.Replace("{$tokenName}", $tokenValue)
        }
        return $result
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $map = [ordered]@{}
        foreach ($key in $Value.Keys) {
            $map[$key] = Apply-Tokens -Value $Value[$key] -Tokens $Tokens
        }
        return $map
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @()
        foreach ($item in $Value) {
            $items += ,(Apply-Tokens -Value $item -Tokens $Tokens)
        }
        return $items
    }

    return $Value
}

function Read-TokenFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not (Test-Path -LiteralPath $resolvedPath)) {
        throw "Token file was not found: $resolvedPath"
    }

    $tokenRoot = ConvertTo-JsonCompatible -Value (ConvertFrom-Json -InputObject (Get-Content -LiteralPath $resolvedPath -Raw))
    $tokenMap = Get-MapValue -Map $tokenRoot -Key "tokens"
    if ($null -eq $tokenMap) {
        $tokenMap = $tokenRoot
    }

    if ($null -eq $tokenMap -or -not ($tokenMap -is [System.Collections.IDictionary])) {
        throw "Token file must contain a JSON object, or an object with a 'tokens' object."
    }

    return $tokenMap
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

function Get-SpSearchControlsFromCanvas {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Canvas
    )

    $controls = @()
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

        $controls += ,[ordered]@{
            component = $componentName
            title = $title
            componentId = $componentId
            instanceId = $instanceId
            occurrence = $countsByComponent[$componentName]
            searchContextId = Get-MapValue -Map $properties -Key "searchContextId"
            control = $control
            webPartData = $webPartData
            properties = $properties
        }
    }

    return $controls
}

function Get-PropertyChangeSummary {
    param(
        [AllowNull()]
        [object]$Before,

        [AllowNull()]
        [object]$After
    )

    $beforeKeys = @()
    $afterKeys = @()

    if ($Before -is [System.Collections.IDictionary]) {
        $beforeKeys = @($Before.Keys)
    }
    if ($After -is [System.Collections.IDictionary]) {
        $afterKeys = @($After.Keys)
    }

    $allKeys = @($beforeKeys + $afterKeys | Sort-Object -Unique)
    $changedKeys = @()

    foreach ($key in $allKeys) {
        $beforeValue = if ($Before -is [System.Collections.IDictionary] -and $Before.Contains($key)) { $Before[$key] } else { $null }
        $afterValue = if ($After -is [System.Collections.IDictionary] -and $After.Contains($key)) { $After[$key] } else { $null }
        $beforeJson = $beforeValue | ConvertTo-Json -Depth 100 -Compress
        $afterJson = $afterValue | ConvertTo-Json -Depth 100 -Compress
        if ($beforeJson -ne $afterJson) {
            $changedKeys += $key
        }
    }

    return [ordered]@{
        Count = $changedKeys.Count
        Keys = $changedKeys
    }
}

function Find-MatchingTargetControl {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$ExportedWebPart,

        [Parameter(Mandatory = $true)]
        [array]$TargetControls,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$UsedTargetIds
    )

    $exportedComponent = Normalize-SpSearchComponentName -Name (Get-MapValue -Map $ExportedWebPart -Key "component") -ComponentId (Get-MapValue -Map $ExportedWebPart -Key "componentId")
    $exportedInstanceId = Get-MapValue -Map $ExportedWebPart -Key "instanceId"
    $exportedOccurrence = Get-MapValue -Map $ExportedWebPart -Key "occurrence"
    $exportedSearchContextId = Get-MapValue -Map $ExportedWebPart -Key "searchContextId"

    if ([string]::IsNullOrWhiteSpace($exportedSearchContextId)) {
        $exportedProperties = Get-MapValue -Map $ExportedWebPart -Key "properties"
        $exportedSearchContextId = Get-MapValue -Map $exportedProperties -Key "searchContextId"
    }

    if (-not $exportedComponent) {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($exportedInstanceId)) {
        $instanceMatch = $TargetControls | Where-Object {
            $_.instanceId -eq $exportedInstanceId -and
            $_.component -eq $exportedComponent -and
            -not $UsedTargetIds.Contains($_.instanceId)
        } | Select-Object -First 1
        if ($instanceMatch) {
            return $instanceMatch
        }
    }

    $componentCandidates = @($TargetControls | Where-Object {
        $_.component -eq $exportedComponent -and
        -not $UsedTargetIds.Contains($_.instanceId)
    })

    if (-not [string]::IsNullOrWhiteSpace($exportedSearchContextId)) {
        $contextOccurrenceMatch = $componentCandidates | Where-Object {
            $_.searchContextId -eq $exportedSearchContextId -and
            $_.occurrence -eq $exportedOccurrence
        } | Select-Object -First 1
        if ($contextOccurrenceMatch) {
            return $contextOccurrenceMatch
        }
    }

    $occurrenceMatch = $componentCandidates | Where-Object {
        $_.occurrence -eq $exportedOccurrence
    } | Select-Object -First 1
    if ($occurrenceMatch) {
        return $occurrenceMatch
    }

    if (-not [string]::IsNullOrWhiteSpace($exportedSearchContextId)) {
        $contextMatches = @($componentCandidates | Where-Object {
            $_.searchContextId -eq $exportedSearchContextId
        })
        if ($contextMatches.Count -eq 1) {
            return $contextMatches[0]
        }
    }

    return $null
}

$normalizedSiteUrl = Normalize-SiteUrl -Value $SiteUrl
$leafName = Normalize-PageLeafName -Value $PageName
$pageNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($leafName)
$resolvedConfigPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ConfigPath)

if (-not (Test-Path -LiteralPath $resolvedConfigPath)) {
    throw "Config file was not found: $resolvedConfigPath"
}

$config = ConvertTo-JsonCompatible -Value (ConvertFrom-Json -InputObject (Get-Content -LiteralPath $resolvedConfigPath -Raw))
$webPartsToImport = @(Get-MapValue -Map $config -Key "webParts")
if ($webPartsToImport.Count -eq 0) {
    throw "Config file does not contain any webParts entries."
}

$tokens = [ordered]@{
    siteUrl = $normalizedSiteUrl
}
if (-not [string]::IsNullOrWhiteSpace($TokenFile)) {
    $fileTokens = Read-TokenFile -Path $TokenFile
    foreach ($key in $fileTokens.Keys) {
        $tokens[$key] = $fileTokens[$key]
    }
}

Write-Host ""
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host " SP Search - Import Page Configuration" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Site:   $normalizedSiteUrl" -ForegroundColor White
Write-Host "  Page:   $leafName" -ForegroundColor White
Write-Host "  Config: $resolvedConfigPath" -ForegroundColor White
Write-Host "  Publish after import: $Publish" -ForegroundColor White
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
$targetControls = @(Get-SpSearchControlsFromCanvas -Canvas $canvas)
if ($targetControls.Count -eq 0) {
    throw "No SP Search web parts were found on target page '$leafName'. Provision the page first, then import the configuration."
}

$usedTargetIds = @{}
$matched = @()
$missing = @()

foreach ($exportedWebPart in $webPartsToImport) {
    $targetControl = Find-MatchingTargetControl -ExportedWebPart $exportedWebPart -TargetControls $targetControls -UsedTargetIds $usedTargetIds
    if (-not $targetControl) {
        $missing += $exportedWebPart
        continue
    }

    $newProperties = Apply-Tokens -Value (Get-MapValue -Map $exportedWebPart -Key "properties") -Tokens $tokens
    if ($null -eq $newProperties) {
        $newProperties = [ordered]@{}
    }

    $changeSummary = Get-PropertyChangeSummary -Before $targetControl.properties -After $newProperties
    $targetControl.webPartData["properties"] = $newProperties

    $newServerProcessedContent = Get-MapValue -Map $exportedWebPart -Key "serverProcessedContent"
    if ($null -ne $newServerProcessedContent) {
        $targetControl.webPartData["serverProcessedContent"] = Apply-Tokens -Value $newServerProcessedContent -Tokens $tokens
    }

    $newDynamicDataPaths = Get-MapValue -Map $exportedWebPart -Key "dynamicDataPaths"
    if ($null -ne $newDynamicDataPaths) {
        $targetControl.webPartData["dynamicDataPaths"] = Apply-Tokens -Value $newDynamicDataPaths -Tokens $tokens
    }

    $newDynamicDataValues = Get-MapValue -Map $exportedWebPart -Key "dynamicDataValues"
    if ($null -ne $newDynamicDataValues) {
        $targetControl.webPartData["dynamicDataValues"] = Apply-Tokens -Value $newDynamicDataValues -Tokens $tokens
    }

    $usedTargetIds[$targetControl.instanceId] = $true
    $matched += ,[ordered]@{
        Exported = $exportedWebPart
        Target = $targetControl
        ChangedKeys = $changeSummary.Keys
        ChangedCount = $changeSummary.Count
    }
}

Write-Host ""
Write-Host "Match summary:" -ForegroundColor Cyan
foreach ($entry in $matched) {
    $component = $entry.Target.component
    $context = $entry.Target.searchContextId
    if ([string]::IsNullOrWhiteSpace($context)) {
        $context = "(no context)"
    }

    $changeText = "no property changes"
    if ($entry.ChangedCount -gt 0) {
        $previewKeys = @($entry.ChangedKeys | Select-Object -First 8)
        $suffix = ""
        if ($entry.ChangedCount -gt $previewKeys.Count) {
            $suffix = ", ..."
        }
        $changeText = "$($entry.ChangedCount) changed key(s): $($previewKeys -join ', ')$suffix"
    }

    Write-Host ("  - {0} occurrence {1}, context {2}: {3}" -f $component, $entry.Target.occurrence, $context, $changeText) -ForegroundColor White
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "Missing target web parts:" -ForegroundColor Yellow
    foreach ($webPart in $missing) {
        $component = Get-MapValue -Map $webPart -Key "component"
        $occurrence = Get-MapValue -Map $webPart -Key "occurrence"
        $context = Get-MapValue -Map $webPart -Key "searchContextId"
        if ([string]::IsNullOrWhiteSpace($context)) {
            $context = "(no context)"
        }
        Write-Host ("  - {0} occurrence {1}, context {2}" -f $component, $occurrence, $context) -ForegroundColor Yellow
    }
    throw "Import stopped because $($missing.Count) exported SP Search web part(s) could not be matched on target page '$leafName'."
}

$newCanvasJson = $canvas | ConvertTo-Json -Depth 100 -Compress
$targetDescription = "$leafName on $normalizedSiteUrl"
$shouldUpdatePage = $false
if ($WhatIfPreference) {
    $shouldUpdatePage = $PSCmdlet.ShouldProcess($targetDescription, "Update SP Search web part properties in CanvasContent1")
} elseif ($Force) {
    $shouldUpdatePage = $true
} else {
    $shouldUpdatePage = $PSCmdlet.ShouldProcess($targetDescription, "Update SP Search web part properties in CanvasContent1")
}

if ($shouldUpdatePage) {
    Set-PnPListItem -List $pageInfo.ListName -Identity $item.Id -Values @{ CanvasContent1 = $newCanvasJson } -ErrorAction Stop | Out-Null
    Write-Host ""
    Write-Host "Updated CanvasContent1 for $leafName" -ForegroundColor Green

    if ($Publish) {
        $shouldPublishPage = $false
        if ($WhatIfPreference) {
            $shouldPublishPage = $PSCmdlet.ShouldProcess($targetDescription, "Publish page")
        } elseif ($Force) {
            $shouldPublishPage = $true
        } else {
            $shouldPublishPage = $PSCmdlet.ShouldProcess($targetDescription, "Publish page")
        }

        if ($shouldPublishPage) {
            Set-PnPPage -Identity $pageNameWithoutExtension -Publish -ErrorAction Stop | Out-Null
            Write-Host "Published $leafName" -ForegroundColor Green
        }
    }
} else {
    Write-Host ""
    Write-Host "No changes were written." -ForegroundColor Yellow
}
