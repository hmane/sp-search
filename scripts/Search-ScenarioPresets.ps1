<#
.SYNOPSIS
    SP Search scenario preset helper — retrieve preset definitions and deploy
    a pre-configured search page for a named scenario.

.DESCRIPTION
    Mirrors the TypeScript SCENARIO_PRESETS registry so PowerShell-driven
    provisioning can wire up all five web parts consistently with the same
    defaults that the property pane picker applies in the browser.

    Two public entry points:
      Get-SearchScenarioPreset   — returns a hashtable describing a preset
      Invoke-SearchScenarioPage  — creates (or updates) a SharePoint page
                                   with all five web parts pre-configured

.EXAMPLE
    # Print the documents preset
    $p = Get-SearchScenarioPreset -Name "documents"
    $p | ConvertTo-Json -Depth 5

.EXAMPLE
    # Deploy a news search page
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/intranet" -Interactive
    Invoke-SearchScenarioPage -SiteUrl "https://contoso.sharepoint.com/sites/intranet" `
                              -PageName "news-search" `
                              -ScenarioName "news" `
                              -SearchContextId "news-ctx" `
                              -PageTitle "News Search"
#>

#Requires -Modules PnP.PowerShell

# ── Preset registry ────────────────────────────────────────────────────────────
# Each hashtable mirrors the IScenarioPreset TypeScript interface.
# Keep in sync with src/webparts/spSearchResults/presets/searchPresets.ts

$PRESET_REGISTRY = @{

  general = @{
    id              = 'general'
    label           = 'General'
    description     = 'Broad search across all content. List layout with flexible refiners.'
    queryTemplate   = '{searchTerms}'
    defaultLayout   = 'list'
    layouts         = @{ list=$true; compact=$true; grid=$true; card=$false; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';            alias='Title'    }
      @{ property='Author';           alias='Author'   }
      @{ property='LastModifiedTime'; alias='Modified' }
      @{ property='FileType';         alias='Type'     }
      @{ property='FileSize';         alias='Size'     }
      @{ property='Path';             alias='URL'      }
      @{ property='SiteName';         alias='Site'     }
    )
    sortableProperties = @(
      @{ property='LastModifiedTime'; label='Date Modified'; direction='Descending' }
      @{ property='Title';            label='Title';         direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='FileType';         label='File type';     urlAlias='ft'; filterType='checkbox' }
      @{ managedProperty='LastModifiedTime'; label='Modified date'; urlAlias='md'; filterType='daterange' }
      @{ managedProperty='AuthorOWSUSER';    label='Author';        urlAlias='au'; filterType='people' }
    )
    verticalSuggestions = @()
  }

  documents = @{
    id              = 'documents'
    label           = 'Documents'
    description     = 'Scoped to SharePoint documents. Compact and grid layouts, file-type and date refiners.'
    queryTemplate   = '{searchTerms} IsDocument:1'
    defaultLayout   = 'list'
    layouts         = @{ list=$true; compact=$true; grid=$true; card=$false; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';                 alias='Title'    }
      @{ property='Author';                alias='Author'   }
      @{ property='LastModifiedTime';      alias='Modified' }
      @{ property='FileType';              alias='Type'     }
      @{ property='FileSize';              alias='Size'     }
      @{ property='Path';                  alias='URL'      }
      @{ property='SiteName';              alias='Site'     }
      @{ property='HitHighlightedSummary'; alias='Summary'  }
    )
    sortableProperties = @(
      @{ property='LastModifiedTime'; label='Date Modified'; direction='Descending' }
      @{ property='Title';            label='Title';         direction='Ascending'  }
      @{ property='FileSize';         label='File Size';     direction='Descending' }
      @{ property='Author';           label='Author';        direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='FileType';         label='File type';     urlAlias='ft'; filterType='checkbox' }
      @{ managedProperty='LastModifiedTime'; label='Modified date'; urlAlias='md'; filterType='daterange' }
      @{ managedProperty='AuthorOWSUSER';    label='Author';        urlAlias='au'; filterType='people'   }
      @{ managedProperty='SiteName';         label='Site';          urlAlias='si'; filterType='checkbox' }
    )
    verticalSuggestions = @(
      @{ key='all';  label='All Documents'; queryTemplate='{searchTerms} IsDocument:1' }
      @{ key='docx'; label='Word';          queryTemplate='{searchTerms} IsDocument:1 FileType:docx' }
      @{ key='xlsx'; label='Excel';         queryTemplate='{searchTerms} IsDocument:1 FileType:xlsx' }
      @{ key='pptx'; label='PowerPoint';    queryTemplate='{searchTerms} IsDocument:1 FileType:pptx' }
      @{ key='pdf';  label='PDF';           queryTemplate='{searchTerms} IsDocument:1 FileType:pdf'  }
    )
  }

  people = @{
    id              = 'people'
    label           = 'People'
    description     = 'People directory via Microsoft Graph. Requires graph-people provider in a vertical.'
    queryTemplate   = '{searchTerms}'
    defaultLayout   = 'people'
    layouts         = @{ list=$false; compact=$false; grid=$false; card=$false; people=$true; gallery=$false }
    selectedProperties = @(
      @{ property='Title';        alias='Name'       }
      @{ property='WorkEmail';    alias='Email'      }
      @{ property='JobTitle';     alias='Job Title'  }
      @{ property='Department';   alias='Department' }
      @{ property='OfficeNumber'; alias='Office'     }
      @{ property='WorkPhone';    alias='Phone'      }
      @{ property='SPS-Skills';   alias='Skills'     }
      @{ property='AboutMe';      alias='About Me'   }
      @{ property='PictureURL';   alias='Photo'      }
    )
    sortableProperties = @(
      @{ property='Title';      label='Name';       direction='Ascending' }
      @{ property='Department'; label='Department'; direction='Ascending' }
    )
    dataProviderHint   = 'graph-people'
    filterSuggestions  = @(
      @{ managedProperty='Department';   label='Department'; urlAlias='dp'; filterType='checkbox' }
      @{ managedProperty='JobTitle';     label='Job title';  urlAlias='jt'; filterType='checkbox' }
      @{ managedProperty='OfficeNumber'; label='Office';     urlAlias='of'; filterType='checkbox' }
    )
    verticalSuggestions = @(
      @{ key='people'; label='People'; dataProvider='graph-people' }
    )
  }

  news = @{
    id              = 'news'
    label           = 'News'
    description     = 'SharePoint News pages (PromotedState:2). Card layout, sorted by published date.'
    queryTemplate   = '{searchTerms} PromotedState:2'
    defaultLayout   = 'card'
    layouts         = @{ list=$true; compact=$false; grid=$false; card=$true; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';                 alias='Title'       }
      @{ property='Author';                alias='Author'      }
      @{ property='Created';               alias='Published'   }
      @{ property='PictureThumbnailURL';   alias='Thumbnail'   }
      @{ property='HitHighlightedSummary'; alias='Description' }
      @{ property='SiteName';              alias='Site'        }
      @{ property='Path';                  alias='URL'         }
    )
    sortableProperties = @(
      @{ property='Created'; label='Published'; direction='Descending' }
      @{ property='Title';   label='Title';     direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='Created';       label='Published date'; urlAlias='pd'; filterType='daterange' }
      @{ managedProperty='AuthorOWSUSER'; label='Author';         urlAlias='au'; filterType='people'   }
      @{ managedProperty='SiteName';      label='Site';           urlAlias='si'; filterType='checkbox' }
    )
    verticalSuggestions = @(
      @{ key='news'; label='News'; queryTemplate='{searchTerms} PromotedState:2' }
    )
  }

  media = @{
    id              = 'media'
    label           = 'Media'
    description     = 'Images and video files. Gallery layout with thumbnail optimisation.'
    queryTemplate   = '{searchTerms} (FileType:jpg OR FileType:jpeg OR FileType:png OR FileType:gif OR FileType:mp4 OR FileType:mov)'
    defaultLayout   = 'gallery'
    layouts         = @{ list=$false; compact=$false; grid=$false; card=$true; people=$false; gallery=$true }
    selectedProperties = @(
      @{ property='Title';               alias='Title'     }
      @{ property='PictureThumbnailURL'; alias='Thumbnail' }
      @{ property='FileType';            alias='Type'      }
      @{ property='FileSize';            alias='Size'      }
      @{ property='LastModifiedTime';    alias='Modified'  }
      @{ property='Author';              alias='Author'    }
      @{ property='SiteName';            alias='Site'      }
      @{ property='Path';                alias='URL'       }
    )
    sortableProperties = @(
      @{ property='LastModifiedTime'; label='Date Modified'; direction='Descending' }
      @{ property='Title';            label='Title';         direction='Ascending'  }
      @{ property='FileSize';         label='File Size';     direction='Descending' }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='FileType';         label='File type';     urlAlias='ft'; filterType='checkbox' }
      @{ managedProperty='LastModifiedTime'; label='Modified date'; urlAlias='md'; filterType='daterange' }
      @{ managedProperty='SiteName';         label='Site';          urlAlias='si'; filterType='checkbox' }
    )
    verticalSuggestions = @(
      @{ key='images'; label='Images'; queryTemplate='{searchTerms} (FileType:jpg OR FileType:jpeg OR FileType:png OR FileType:gif)' }
      @{ key='video';  label='Video';  queryTemplate='{searchTerms} (FileType:mp4 OR FileType:mov)' }
    )
  }

  'hub-search' = @{
    id              = 'hub-search'
    label           = 'Hub Search'
    description     = 'Cross-site intranet search scoped to a hub. All content types with document, news, and page verticals.'
    queryTemplate   = '{searchTerms}'
    defaultLayout   = 'list'
    layouts         = @{ list=$true; compact=$true; grid=$true; card=$true; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';                 alias='Title'    }
      @{ property='Author';                alias='Author'   }
      @{ property='LastModifiedTime';      alias='Modified' }
      @{ property='FileType';              alias='Type'     }
      @{ property='SiteName';              alias='Site'     }
      @{ property='HitHighlightedSummary'; alias='Summary'  }
      @{ property='Path';                  alias='URL'      }
    )
    sortableProperties = @(
      @{ property='LastModifiedTime'; label='Date Modified'; direction='Descending' }
      @{ property='Title';            label='Title';         direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='FileType';         label='Content type';  urlAlias='ft'; filterType='checkbox'  }
      @{ managedProperty='SiteName';         label='Site';          urlAlias='si'; filterType='checkbox'  }
      @{ managedProperty='AuthorOWSUSER';    label='Author';        urlAlias='au'; filterType='people'    }
      @{ managedProperty='LastModifiedTime'; label='Modified date'; urlAlias='md'; filterType='daterange' }
    )
    verticalSuggestions = @(
      @{ key='all';       label='All';       queryTemplate='{searchTerms}'                          }
      @{ key='documents'; label='Documents'; queryTemplate='{searchTerms} IsDocument:1'             }
      @{ key='news';      label='News';      queryTemplate='{searchTerms} PromotedState:2'          }
      @{ key='pages';     label='Pages';     queryTemplate='{searchTerms} contentclass:STS_ListItem_Pages -PromotedState:2' }
    )
  }

  'knowledge-base' = @{
    id              = 'knowledge-base'
    label           = 'Knowledge Base'
    description     = 'Knowledge articles, how-to guides, and reference documents. Card layout with rich preview.'
    queryTemplate   = '{searchTerms} (IsDocument:1 OR contentclass:STS_ListItem_Pages)'
    defaultLayout   = 'card'
    layouts         = @{ list=$true; compact=$false; grid=$true; card=$true; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';                 alias='Title'     }
      @{ property='Author';                alias='Author'    }
      @{ property='Created';               alias='Published' }
      @{ property='HitHighlightedSummary'; alias='Summary'   }
      @{ property='ContentType';           alias='Category'  }
      @{ property='SiteName';              alias='Site'      }
      @{ property='PictureThumbnailURL';   alias='Thumbnail' }
      @{ property='Path';                  alias='URL'       }
    )
    sortableProperties = @(
      @{ property='Created';          label='Published';    direction='Descending' }
      @{ property='LastModifiedTime'; label='Last Updated'; direction='Descending' }
      @{ property='Title';            label='Title';        direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='ContentType';   label='Category';       urlAlias='ct'; filterType='checkbox'  }
      @{ managedProperty='SiteName';      label='Site';           urlAlias='si'; filterType='checkbox'  }
      @{ managedProperty='AuthorOWSUSER'; label='Author';         urlAlias='au'; filterType='people'    }
      @{ managedProperty='Created';       label='Published date'; urlAlias='pd'; filterType='daterange' }
    )
    verticalSuggestions = @(
      @{ key='all';        label='All';        queryTemplate='{searchTerms} (IsDocument:1 OR contentclass:STS_ListItem_Pages)' }
      @{ key='articles';   label='Articles';   queryTemplate='{searchTerms} contentclass:STS_ListItem_Pages'                   }
      @{ key='guides';     label='Guides';     queryTemplate='{searchTerms} IsDocument:1 (FileType:pdf OR FileType:docx)'      }
    )
  }

  'policy-search' = @{
    id              = 'policy-search'
    label           = 'Policy Search'
    description     = 'Corporate policies, procedures, and compliance documents. Scoped to PDF and Office files.'
    queryTemplate   = '{searchTerms} IsDocument:1 (FileType:pdf OR FileType:docx OR FileType:doc OR FileType:xlsx)'
    defaultLayout   = 'list'
    layouts         = @{ list=$true; compact=$true; grid=$true; card=$false; people=$false; gallery=$false }
    selectedProperties = @(
      @{ property='Title';                 alias='Title'    }
      @{ property='Author';                alias='Owner'    }
      @{ property='LastModifiedTime';      alias='Reviewed' }
      @{ property='FileType';              alias='Type'     }
      @{ property='FileSize';              alias='Size'     }
      @{ property='SiteName';              alias='Source'   }
      @{ property='HitHighlightedSummary'; alias='Summary'  }
      @{ property='Path';                  alias='URL'      }
    )
    sortableProperties = @(
      @{ property='Title';            label='Title';         direction='Ascending'  }
      @{ property='LastModifiedTime'; label='Last Reviewed'; direction='Descending' }
      @{ property='Author';           label='Owner';         direction='Ascending'  }
    )
    dataProviderHint   = 'sharepoint-search'
    filterSuggestions  = @(
      @{ managedProperty='FileType';         label='File type';     urlAlias='ft'; filterType='checkbox'  }
      @{ managedProperty='SiteName';         label='Source';        urlAlias='si'; filterType='checkbox'  }
      @{ managedProperty='AuthorOWSUSER';    label='Policy owner';  urlAlias='au'; filterType='people'    }
      @{ managedProperty='LastModifiedTime'; label='Last reviewed'; urlAlias='md'; filterType='daterange' }
    )
    verticalSuggestions = @(
      @{ key='all';        label='All';        queryTemplate='{searchTerms} IsDocument:1 (FileType:pdf OR FileType:docx OR FileType:doc OR FileType:xlsx)' }
      @{ key='pdf';        label='PDF';        queryTemplate='{searchTerms} IsDocument:1 FileType:pdf'                                                      }
      @{ key='word';       label='Word';       queryTemplate='{searchTerms} IsDocument:1 (FileType:docx OR FileType:doc)'                                   }
    )
  }
}

# ── Public functions ───────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Returns the scenario preset hashtable for a given preset name.

.PARAMETER Name
    Preset ID: general | documents | people | news | media | hub-search | knowledge-base | policy-search

.OUTPUTS
    Hashtable — mirrors IScenarioPreset from searchPresets.ts, or $null if
    the name is unknown.

.EXAMPLE
    $p = Get-SearchScenarioPreset -Name "documents"
    Write-Host "Query template: $($p.queryTemplate)"
    Write-Host "Suggested filters:"
    $p.filterSuggestions | ForEach-Object { "  $($_.label) ($($_.filterType))" }
#>
function Get-SearchScenarioPreset {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Name
  )

  $preset = $PRESET_REGISTRY[$Name.ToLower()]
  if (-not $preset) {
    Write-Warning "Unknown scenario preset '$Name'. Valid names: $($PRESET_REGISTRY.Keys -join ', ')"
    return $null
  }
  return $preset
}

<#
.SYNOPSIS
    Creates or updates a SharePoint page pre-configured with all SP Search
    web parts for the named scenario.

.DESCRIPTION
    Provisions a modern SharePoint page containing:
      - SP Search Box
      - SP Search Verticals (if the preset has verticalSuggestions)
      - SP Search Filters (sidebar)
      - SP Search Results (pre-configured with preset layout, query template,
        and selected properties)
      - SP Search Manager

    All web parts share the same SearchContextId so they communicate via the
    Zustand store.

.PARAMETER SiteUrl
    Absolute URL of the target SharePoint site.

.PARAMETER PageName
    URL slug for the new page, e.g. "documents-search" → /SitePages/documents-search.aspx

.PARAMETER ScenarioName
    Preset ID: general | documents | people | news | media | hub-search | knowledge-base | policy-search

.PARAMETER SearchContextId
    Shared context ID string wired into every web part. Defaults to the preset
    ID (e.g. "documents").

.PARAMETER PageTitle
    Page title. Defaults to "<Label> Search", e.g. "Documents Search".

.PARAMETER OverwriteExisting
    If $true and the page already exists, overwrite it. Defaults to $false.

.EXAMPLE
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/intranet" -Interactive

    Invoke-SearchScenarioPage `
      -SiteUrl       "https://contoso.sharepoint.com/sites/intranet" `
      -PageName      "documents-search" `
      -ScenarioName  "documents" `
      -PageTitle     "Document Search"
#>
function Invoke-SearchScenarioPage {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$PageName,
    [Parameter(Mandatory)][string]$ScenarioName,
    [string]$SearchContextId,
    [string]$PageTitle,
    [switch]$OverwriteExisting
  )

  # ── Resolve preset ──────────────────────────────────────────────────────────
  $preset = Get-SearchScenarioPreset -Name $ScenarioName
  if (-not $preset) { return }

  if (-not $SearchContextId) { $SearchContextId = $preset.id }
  if (-not $PageTitle)        { $PageTitle = "$($preset.label) Search" }

  Write-Host "Deploying '$PageTitle' page (preset: $ScenarioName, contextId: $SearchContextId)" -ForegroundColor Cyan

  # ── Ensure connection ───────────────────────────────────────────────────────
  try {
    $null = Get-PnPWeb -ErrorAction Stop
  }
  catch {
    Write-Host "Connecting to $SiteUrl..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -Interactive
  }

  # ── Create / get page ───────────────────────────────────────────────────────
  $pagePath = "$PageName.aspx"
  $existingPage = $null
  try { $existingPage = Get-PnPPage -Identity $pagePath -ErrorAction SilentlyContinue } catch {}

  if ($existingPage -and -not $OverwriteExisting) {
    Write-Warning "Page '$pagePath' already exists. Use -OverwriteExisting to replace it."
    return
  }

  if (-not $PSCmdlet.ShouldProcess($pagePath, "Create/Update search page")) { return }

  if ($existingPage -and $OverwriteExisting) {
    Write-Host "Removing existing page '$pagePath'..." -ForegroundColor Yellow
    Remove-PnPPage -Identity $pagePath -Force
  }

  Write-Host "Creating page '$pagePath'..." -ForegroundColor White
  $page = Add-PnPPage -Name $pagePath -Title $PageTitle -LayoutType Home -CommentsEnabled:$false

  # ── Build web part property JSON ─────────────────────────────────────────────

  # Search Box
  $boxProps = @{
    searchContextId = $SearchContextId
    placeholder     = "Search $($preset.label.ToLower())..."
  } | ConvertTo-Json -Compress

  # Search Results — core preset properties
  $resultsProps = @{
    searchContextId            = $SearchContextId
    defaultLayout              = $preset.defaultLayout
    showListLayout             = $preset.layouts.list
    showCompactLayout          = $preset.layouts.compact
    showGridLayout             = $preset.layouts.grid
    showCardLayout             = $preset.layouts.card
    showPeopleLayout           = $preset.layouts.people
    showGalleryLayout          = $preset.layouts.gallery
    queryTemplate              = $preset.queryTemplate
    showResultCount            = $true
    showSortDropdown           = $true
    pageSize                   = 10
    selectedPropertiesCollection = @(
      $preset.selectedProperties | ForEach-Object -Begin { $idx=0 } -Process {
        @{ uniqueId="preset-sp-$idx"; property=$_.property; alias=$_.alias }
        $idx++
      }
    )
    sortablePropertiesCollection = @(
      $preset.sortableProperties | ForEach-Object -Begin { $idx=0 } -Process {
        @{ uniqueId="preset-sort-$idx"; property=$_.property; label=$_.label; direction=$_.direction }
        $idx++
      }
    )
  } | ConvertTo-Json -Compress -Depth 5

  # Search Filters — add suggested filters as pre-configured filter groups
  $filterGroups = @(
    $preset.filterSuggestions | ForEach-Object -Begin { $idx=0 } -Process {
      @{
        uniqueId       = "preset-filter-$idx"
        displayName    = $_.label
        managedProperty = $_.managedProperty
        urlAlias       = $_.urlAlias
        filterType     = $_.filterType
        isCollapsed    = $false
        showCount      = $true
      }
      $idx++
    }
  )
  $filtersProps = @{
    searchContextId   = $SearchContextId
    filterCollection  = $filterGroups
    showFilterCount   = $true
    filterHeaderLabel = "Filter $($preset.label.ToLower())"
  } | ConvertTo-Json -Compress -Depth 5

  # Search Verticals — add suggested verticals if defined
  $verticalItems = @()
  if ($preset.verticalSuggestions -and $preset.verticalSuggestions.Count -gt 0) {
    $verticalItems = @(
      $preset.verticalSuggestions | ForEach-Object -Begin { $idx=0 } -Process {
        $vItem = @{
          uniqueId      = "preset-vert-$idx"
          tabName       = $_.label
          verticalKey   = $_.key
          queryTemplate = if ($_.queryTemplate) { $_.queryTemplate } else { $preset.queryTemplate }
        }
        if ($_.dataProvider) { $vItem.dataProvider = $_.dataProvider }
        $vItem
        $idx++
      }
    )
  }
  $verticalsProps = @{
    searchContextId    = $SearchContextId
    showBadgeCounts    = $true
    verticalCollection = $verticalItems
  } | ConvertTo-Json -Compress -Depth 5

  # Search Manager
  $managerProps = @{
    searchContextId = $SearchContextId
    displayMode     = 'panel'
  } | ConvertTo-Json -Compress

  # ── Add web parts to page ───────────────────────────────────────────────────
  # Row 1: Search box (full width)
  $page | Add-PnPPageSection -SectionTemplate OneColumn -Order 1

  # Row 2: Filters (left 1/3) + Results (right 2/3)
  $page | Add-PnPPageSection -SectionTemplate TwoColumnLeft -Order 2

  # Row 3: Manager (full width, hidden visual footprint)
  $page | Add-PnPPageSection -SectionTemplate OneColumn -Order 3

  # Add web parts — component names must match installed SPFx package manifest titles
  Write-Host "Adding SP Search Box..." -ForegroundColor White
  Add-PnPPageWebPart -Page $page -Component "SP Search Box" `
    -Section 1 -Column 1 -WebPartProperties $boxProps -ErrorAction SilentlyContinue

  Write-Host "Adding SP Search Results..." -ForegroundColor White
  Add-PnPPageWebPart -Page $page -Component "SP Search Results" `
    -Section 2 -Column 2 -WebPartProperties $resultsProps -ErrorAction SilentlyContinue

  Write-Host "Adding SP Search Filters..." -ForegroundColor White
  Add-PnPPageWebPart -Page $page -Component "SP Search Filters" `
    -Section 2 -Column 1 -WebPartProperties $filtersProps -ErrorAction SilentlyContinue

  if ($verticalItems.Count -gt 0) {
    Write-Host "Adding SP Search Verticals ($($verticalItems.Count) tabs)..." -ForegroundColor White
    # Insert verticals between box and filters/results by adjusting order
    Add-PnPPageWebPart -Page $page -Component "SP Search Verticals" `
      -Section 1 -Column 1 -Order 2 -WebPartProperties $verticalsProps -ErrorAction SilentlyContinue
  }

  Write-Host "Adding SP Search Manager..." -ForegroundColor White
  Add-PnPPageWebPart -Page $page -Component "SP Search Manager" `
    -Section 3 -Column 1 -WebPartProperties $managerProps -ErrorAction SilentlyContinue

  # ── Publish ─────────────────────────────────────────────────────────────────
  Publish-PnPPage -Identity $pagePath
  Write-Host "Page published: $SiteUrl/SitePages/$pagePath" -ForegroundColor Green

  # ── Print configuration hints ───────────────────────────────────────────────
  Write-Host ""
  Write-Host "── Post-deployment configuration hints ──────────────────────────" -ForegroundColor Cyan
  Write-Host "  Preset : $($preset.label)"
  Write-Host "  Context: $SearchContextId"
  Write-Host ""
  Write-Host "  Query template (Results web part):"
  Write-Host "    $($preset.queryTemplate)" -ForegroundColor Yellow
  if ($preset.dataProviderHint -ne 'sharepoint-search') {
    Write-Host ""
    Write-Host "  Data provider: '$($preset.dataProviderHint)' required." -ForegroundColor Yellow
    Write-Host "  Configure via the Verticals web part per-vertical Data Provider dropdown."
  }
  if ($preset.filterSuggestions.Count -gt 0) {
    Write-Host ""
    Write-Host "  Suggested Filters web part configuration:"
    $preset.filterSuggestions | ForEach-Object {
      Write-Host "    - $($_.label) ($($_.managedProperty), type: $($_.filterType))" -ForegroundColor Gray
    }
  }
  Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
}

# ── List all presets ───────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Lists all available scenario presets with their IDs, labels, and descriptions.
#>
function Get-SearchScenarioPresetList {
  [CmdletBinding()]
  param()

  Write-Host ""
  Write-Host "Available SP Search scenario presets:" -ForegroundColor Cyan
  Write-Host ""
  foreach ($key in $PRESET_REGISTRY.Keys | Sort-Object) {
    $p = $PRESET_REGISTRY[$key]
    Write-Host "  $($p.id.PadRight(12)) $($p.label.PadRight(12))  $($p.description)" -ForegroundColor White
  }
  Write-Host ""
}
